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
        /// 把当前 Word 实例的身份信息（PID / HWND / 版本 / 文档 / ROT 注册名）
        /// 连同可直接运行的访问代码复制到剪贴板，供大模型定位并操作本实例。
        /// </summary>
        public void BtnCopyInstanceInfo_Click(Office.IRibbonControl control)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;

                int pid;
                string startTime;
                string exePath;
                using (Process proc = Process.GetCurrentProcess())
                {
                    pid = proc.Id;
                    startTime = SafeGet(() => proc.StartTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                    exePath = SafeGet(() => proc.MainModule.FileName);
                }

                string name = EnsureRotRegistration(app, pid);
                long hwnd = GetMainWindowHandle(app);
                string hwndHex = "0x" + hwnd.ToString("X8", CultureInfo.InvariantCulture);

                Word.Document activeDoc = null;
                try { activeDoc = app.ActiveDocument; } catch { }
                string activeDocPath = activeDoc == null ? "(无打开文档)" : SafeGet(() => activeDoc.FullName);

                int docCount;
                try { docCount = app.Documents.Count; } catch { docCount = 0; }

                StringBuilder docList = new StringBuilder();
                try
                {
                    foreach (Word.Document d in app.Documents)
                    {
                        docList.Append("    - ").AppendLine(SafeGet(() => d.FullName));
                    }
                }
                catch { }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("[WordVSTO 当前 Word 实例] 采集时间 " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                sb.AppendLine();
                sb.AppendLine("实例标识");
                sb.AppendLine("  PID            : " + pid);
                sb.AppendLine("  进程启动时间   : " + startTime);
                sb.AppendLine("  主窗口 HWND    : " + hwndHex);
                sb.AppendLine("  WINWORD.EXE    : " + exePath);
                sb.AppendLine("  Word 版本      : " + SafeGet(() => app.Version) + " (Build " + SafeGet(() => app.Build) + ")");
                sb.AppendLine("  Word 程序路径  : " + SafeGet(() => app.Path));
                sb.AppendLine("  用户名         : " + SafeGet(() => app.UserName));
                sb.AppendLine("  ROT 注册名     : !" + name);
                sb.AppendLine();
                sb.AppendLine("文档");
                sb.AppendLine("  当前文档       : " + activeDocPath);
                sb.AppendLine("  打开文档数     : " + docCount);
                sb.Append(docList.ToString());
                sb.AppendLine();
                sb.AppendLine("--- 访问本实例（Python 3 + pywin32；多个 Word 实例也能精确匹配）---");
                sb.AppendLine("import pythoncom, win32com.client");
                sb.AppendLine("name = \"" + name + "\"");
                sb.AppendLine("ctx = pythoncom.CreateBindCtx(0)");
                sb.AppendLine("rot = pythoncom.GetRunningObjectTable()");
                sb.AppendLine("word = None");
                sb.AppendLine("for mk in rot:");
                sb.AppendLine("    try:");
                sb.AppendLine("        display = mk.GetDisplayName(ctx, None)");
                sb.AppendLine("    except pythoncom.com_error:");
                sb.AppendLine("        continue");
                sb.AppendLine("    if name in display:");
                sb.AppendLine("        unk = rot.GetObject(mk)");
                sb.AppendLine("        word = win32com.client.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))");
                sb.AppendLine("        break");
                sb.AppendLine("assert word is not None, \"未找到该实例：请先在 Word 中点击『复制实例信息』\"");
                sb.AppendLine("print(word.ActiveDocument.FullName)");
                sb.AppendLine();
                sb.AppendLine("--- 校验：确认绑定的是 PID " + pid + " 这个进程 ---");
                sb.AppendLine("import win32process");
                sb.AppendLine("assert win32process.GetWindowThreadProcessId(word.ActiveWindow.Hwnd)[1] == " + pid);
                sb.AppendLine();
                if (activeDocPath.IndexOf(":\\", StringComparison.Ordinal) >= 0)
                {
                    sb.AppendLine("--- 备选（当前文档已保存时可用，无需 ROT 名）---");
                    sb.AppendLine("word = win32com.client.GetObject(r\"" + activeDocPath + "\").Application");
                    sb.AppendLine();
                }
                sb.AppendLine("说明：ROT 条目只在本次 Word 进程存活期间有效；Word 关闭后需重新点击该按钮。");

                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "复制实例信息",
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
