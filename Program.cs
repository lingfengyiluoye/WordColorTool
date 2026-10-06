using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace WordColorTool
{
    /// <summary>
    /// 程序入口：启用视觉样式、创建并显示悬浮条、进入消息循环。
    /// 悬浮条默认只是屏幕右侧一条 36px 宽的窄条，鼠标移入才展开成完整面板。
    /// </summary>
    internal static class Program
    {
        private static MainForm _mainForm;

        [STAThread]
        private static void Main()
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += OnThreadException;
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            _mainForm = new MainForm();
            Application.Run(_mainForm);

            if (_mainForm != null)
            {
                _mainForm.Dispose();
                _mainForm = null;
            }
        }

        private static void OnThreadException(object sender, ThreadExceptionEventArgs e)
        {
            Trace.WriteLine("WordColorTool 未处理的 UI 线程异常: " + e.Exception);
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            object ex = e.ExceptionObject;
            Trace.WriteLine("WordColorTool 未处理的异常: " + ex);
        }
    }
}
