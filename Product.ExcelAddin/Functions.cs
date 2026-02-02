using ExcelDna.Integration;
using Microsoft.Win32;
using System;

public static class Functions
{
    private const string KeyPath = @"HKEY_CURRENT_USER\Software\ProductMVP";

    private static bool IsUnlocked()
    {
        var token = Registry.GetValue(KeyPath, "SessionToken", null) as string;
        var expStr = Registry.GetValue(KeyPath, "TokenExpiresUtcTicks", null) as string;

        if (string.IsNullOrWhiteSpace(token)) return false;
        if (!long.TryParse(expStr, out var expTicks)) return false;

        return DateTime.UtcNow.Ticks <= expTicks;
    }

    [ExcelFunction(Description = "Shows whether workbook is opened via launcher")]
    public static string PRODUCT_STATUS()
    {
        return IsUnlocked()
            ? "OK: Opened via Product Launcher"
            : "BLOCKED: Open using Product Launcher";
    }

    // Example of a protected function
    [ExcelFunction(Description = "Example calculation (only works via launcher)")]
    public static object PRODUCT_ADD(double a, double b)
    {
        if (!IsUnlocked())
            return "BLOCKED";

        return a + b;
    }
}
