using System.Text;

namespace MultiReplace;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        // Shift_JIS(CP932) を使うために必要
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}
