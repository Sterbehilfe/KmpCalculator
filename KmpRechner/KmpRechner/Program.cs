using System.Runtime.Versioning;

namespace KmpRechner;

public static class Program
{
    [SupportedOSPlatform("windows")]
    private static void Main()
    {
        Actions.ConvertToXlsx();
    }
}
