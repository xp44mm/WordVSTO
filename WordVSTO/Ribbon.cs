using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordVSTO
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        // 本实例在运行对象表（ROT）中的唯一名字，供外部进程精确绑定。
        private static string rotName;

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out ComTypes.IRunningObjectTable pprot);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        private static extern int CreateItemMoniker(string lpszDelim, string lpszObject, out ComTypes.IMoniker ppmk);

        public Ribbon()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordVSTO.Ribbon.xml");
        }

        #endregion

        #region 功能区回调

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// 在当前光标/选区处插入 mark+时间戳 书签，
        /// 并把「文档路径 + 书签名」复制到剪贴板。
        /// </summary>
        public void BtnInsertDefaultBookmark_Click(Office.IRibbonControl control)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app.ActiveDocument;
                if (doc == null)
                {
                    return;
                }

                string name = NextBookmarkName(doc);
                Word.Range range = app.Selection.Range;
                doc.Bookmarks.Add(name, range);

                string text = "\"" + doc.FullName + "\"" + Environment.NewLine
                            + "书签：\"" + name + "\"";
                Clipboard.SetText(text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "插入默认书签",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「复制当前文档引用信息」：把定位所需的后台信息（采集时间 / PID / ROT 注册名 /
        /// 主窗口 HWND / WINWORD.EXE 路径）写入剪贴板，尾部追加当前文档的名称与完整路径，
        /// 供大模型精确定位并引用这个文档。
        /// </summary>
        public void BtnCopyActiveDocumentInfo_Click(Office.IRibbonControl control)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;

                int pid;
                string exePath;
                using (Process proc = Process.GetCurrentProcess())
                {
                    pid = proc.Id;
                    exePath = SafeGet(() => proc.MainModule.FileName);
                }

                string name = EnsureRotRegistration(app, pid);
                long hwnd = GetMainWindowHandle(app);
                string hwndHex = "0x" + hwnd.ToString("X8", CultureInfo.InvariantCulture);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("[WordVSTO 当前文档引用] 采集时间 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                sb.AppendLine("PID : " + pid);
                sb.AppendLine("ROT : !" + name);
                sb.AppendLine("HWND : " + hwndHex);
                sb.AppendLine("WINWORD.EXE : " + exePath);

                Word.Document doc = null;
                try
                {
                    doc = app.ActiveDocument;
                }
                catch
                {
                    // 一个文档都没打开时 ActiveDocument 会抛 COM 异常，按「无」处理。
                }

                if (doc == null)
                {
                    sb.AppendLine("文档 : (无打开的文档)");
                    sb.AppendLine("路径 : (无)");
                }
                else
                {
                    sb.AppendLine("文档 : " + SafeGet(() => doc.Name));
                    sb.AppendLine("路径 : " + SafeGet(() => string.IsNullOrEmpty(doc.Path) ? "(未保存)" : doc.FullName));
                }

                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "复制当前文档引用信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 帮助器

        /// <summary>
        /// 把当前 Word Application 以 !WordVSTO.Instance.&lt;PID&gt; 注册到 ROT。
        /// flag=0 为弱注册：ROT 不持有强引用，Word 退出后条目自动消失，不会残留僵尸进程。
        /// </summary>
        private static string EnsureRotRegistration(Word.Application app, int pid)
        {
            if (rotName != null)
            {
                return rotName;
            }

            string name = "WordVSTO.Instance." + pid.ToString(CultureInfo.InvariantCulture);

            ComTypes.IRunningObjectTable table;
            int hr = GetRunningObjectTable(0, out table);
            if (hr < 0) { Marshal.ThrowExceptionForHR(hr); }

            ComTypes.IMoniker moniker;
            hr = CreateItemMoniker("!", name, out moniker);
            if (hr < 0) { Marshal.ThrowExceptionForHR(hr); }

            table.Register(0, app, moniker);
            rotName = name;
            return name;
        }

        private static long GetMainWindowHandle(Word.Application app)
        {
            try
            {
                return (long)(uint)app.ActiveWindow.Hwnd;
            }
            catch
            {
                using (Process proc = Process.GetCurrentProcess())
                {
                    return proc.MainWindowHandle.ToInt64();
                }
            }
        }

        private static string SafeGet(Func<string> getter)
        {
            try
            {
                string value = getter();
                return string.IsNullOrEmpty(value) ? "(无)" : value;
            }
            catch
            {
                return "(读取失败)";
            }
        }

        private static string NextBookmarkName(Word.Document doc)
        {
            string baseName = "mark" + DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            string name = baseName;
            int index = 2;
            while (doc.Bookmarks.Exists(name))
            {
                name = baseName + "_" + index.ToString(CultureInfo.InvariantCulture);
                index++;
            }
            return name;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader reader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
