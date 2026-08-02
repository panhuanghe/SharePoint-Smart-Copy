using System.IO;

namespace SharePointSmartCopy.Services;

public static class SiteUrlHelper
{
    // Extracts the site path segment used to prefix default report filenames, e.g.
    // "https://contoso.sharepoint.com/sites/Marketing" -> "Marketing". OneDrive personal sites have
    // no "/sites/" or "/teams/" segment at all — "/personal/{username}/..." instead — so that's
    // checked last, taking just the username segment (up to the next slash) as the site name, same
    // as the other two.
    public static string ExtractSiteName(string? siteUrl)
    {
        if (string.IsNullOrWhiteSpace(siteUrl)) return "";

        var idx = siteUrl.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase);
        var segmentLength = "/sites/".Length;
        if (idx < 0)
        {
            idx = siteUrl.IndexOf("/teams/", StringComparison.OrdinalIgnoreCase);
            segmentLength = "/teams/".Length;
        }
        if (idx < 0)
        {
            idx = siteUrl.IndexOf("/personal/", StringComparison.OrdinalIgnoreCase);
            segmentLength = "/personal/".Length;
        }
        if (idx < 0) return "";

        var rest = siteUrl[(idx + segmentLength)..].TrimStart('/');
        var slashIdx = rest.IndexOf('/');
        var name = slashIdx >= 0 ? rest[..slashIdx] : rest;

        // Non-ASCII site/user names (e.g. "Événements", "会議室") arrive percent-encoded in the raw
        // URL string — without decoding, a localized name would show up in report filenames as an
        // unreadable "%C3%89v%C3%A9nements" escape sequence instead of the real characters.
        try { name = Uri.UnescapeDataString(name); } catch (FormatException) { /* keep encoded form */ }

        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }

    // "Source-Target-" prefix for default report filenames. Omitted entirely if either side's
    // site name can't be determined, rather than producing a lopsided "Source--" or "-Target-".
    public static string ReportFilenamePrefix(string? sourceUrl, string? targetUrl, bool enabled = true)
    {
        if (!enabled) return "";
        var source = ExtractSiteName(sourceUrl);
        var target = ExtractSiteName(targetUrl);
        return source.Length > 0 && target.Length > 0 ? $"{source}-{target}-" : "";
    }
}
