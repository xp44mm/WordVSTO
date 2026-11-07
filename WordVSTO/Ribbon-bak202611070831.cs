using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
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
        /// 把本实例在运行对象表（ROT）中的访问 key 复制到剪贴板。
        /// 剪贴板内容只有一行，形如 !WordVSTO.Instance.31156，不含任何说明文字。
        /// </summary>
        public void BtnCopyInstanceInfo_Click(Office.IRibbonControl control)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                int pid;
                using (Process proc = Process.GetCurrentProcess())
                {
                    pid = proc.Id;
                }

                string name = EnsureRotRegistration(app, pid);
                Clipboard.SetText("!" + name);
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
