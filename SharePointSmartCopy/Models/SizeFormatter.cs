namespace SharePointSmartCopy.Models;

public static class SizeFormatter
{
    public static string FormatBytes(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:N1} KB";
        if (bytes < 1024L * 1024 * 1024) return $"{bytes / (1024.0 * 1024):N1} MB";
        return $"{bytes / (1024.0 * 1024 * 1024):N2} GB";
    }

    public static string FormatBytes(long? bytes) => bytes.HasValue ? FormatBytes(bytes.Value) : string.Empty;
}
