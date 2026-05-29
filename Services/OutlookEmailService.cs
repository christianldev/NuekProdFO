using System.IO;
using Microsoft.Win32;
using NuekProdFO.Models;

namespace NuekProdFO.Services;

/// <summary>
/// Sends email via locally installed Outlook using late-bound COM (no PIA / interop DLLs needed).
/// Works on .NET 5+ where the GAC is not available.
/// </summary>
public class OutlookEmailService
{
    private const int OlMailItem = 0;
    private const int OlCC = 2;
    private const int OlFormatHtml = 2;
    private const string PreferredSignatureName = "nuek_firma";

    public void SendTicketNotification(OutlookEmailSettings settings, TicketRecord record)
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application", throwOnError: true)!;
        dynamic outlookApp = Activator.CreateInstance(outlookType)!;
        dynamic mail;

        try
        {
            mail = PrepareMail(outlookApp, settings, record, includeSignature: true, includePackageLink: true);
        }
        catch (Exception ex) when (IsExpectedRangeError(ex))
        {
            // Primer fallback en preparacion: mantenemos firma, simplificando HTML.
            mail = PrepareMail(outlookApp, settings, record, includeSignature: true, includePackageLink: false);
        }

        TrySendWithFallback(outlookApp, settings, record, mail);
    }

    private static dynamic PrepareMail(
        dynamic outlookApp,
        OutlookEmailSettings settings,
        TicketRecord record,
        bool includeSignature,
        bool includePackageLink)
    {
        dynamic mail = outlookApp.CreateItem(OlMailItem);

        TrySetSenderAccount(outlookApp, mail, settings.SenderAccount);

        mail.BodyFormat = OlFormatHtml;
        mail.Subject = BuildSubject(record);

        var signatureHtml = includeSignature
            ? CaptureSignatureHtml(mail, settings.HideDraftWhenApplyingSignature)
            : string.Empty;

        mail.HTMLBody = BuildBody(record, signatureHtml, includePackageLink);

        AddRecipients(mail, settings);

        var allResolved = (bool)mail.Recipients.ResolveAll();
        if (!allResolved)
        {
            throw new InvalidOperationException("No fue posible resolver uno o mas destinatarios en Outlook.");
        }

        return mail;
    }

    private static void AddRecipients(dynamic mail, OutlookEmailSettings settings)
    {
        foreach (var address in settings.ToRecipients)
        {
            mail.Recipients.Add(address);
        }

        foreach (var address in settings.CcRecipients)
        {
            dynamic recipient = mail.Recipients.Add(address);
            recipient.Type = OlCC;
        }
    }

    private static void TrySetSenderAccount(dynamic outlookApp, dynamic mail, string senderAccount)
    {
        if (string.IsNullOrWhiteSpace(senderAccount))
        {
            return;
        }

        dynamic? account = FindAccount(outlookApp, senderAccount);
        if (account is not null)
        {
            mail.SendUsingAccount = account;
        }
    }

    private static string CaptureSignatureHtml(dynamic mail, bool hideDraft)
    {
        try
        {
            if (hideDraft)
            {
                _ = mail.GetInspector;
            }
            else
            {
                mail.Display(false);
            }

            var htmlFromOutlook = (string)mail.HTMLBody;
            if (!string.IsNullOrWhiteSpace(htmlFromOutlook) && !LooksLikeEmptyBody(htmlFromOutlook))
            {
                return ExtractSignatureFragment(htmlFromOutlook);
            }

            // Fallback: tomamos la firma desde archivos locales de Outlook.
            var preferred = ReadSignatureHtmlByName(PreferredSignatureName);
            if (!string.IsNullOrWhiteSpace(preferred))
            {
                return ExtractSignatureFragment(preferred);
            }

            var defaultSignatureName = ReadDefaultSignatureName();
            return ExtractSignatureFragment(ReadSignatureHtmlByName(defaultSignatureName));
        }
        catch
        {
            var preferred = ReadSignatureHtmlByName(PreferredSignatureName);
            if (!string.IsNullOrWhiteSpace(preferred))
            {
                return ExtractSignatureFragment(preferred);
            }

            var defaultSignatureName = ReadDefaultSignatureName();
            var fallback = ReadSignatureHtmlByName(defaultSignatureName);
            if (!string.IsNullOrWhiteSpace(fallback))
            {
                return ExtractSignatureFragment(fallback);
            }

            return string.Empty;
        }
    }

    private static dynamic? FindAccount(dynamic outlookApp, string emailAddress)
    {
        int count = (int)outlookApp.Session.Accounts.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic account = outlookApp.Session.Accounts[i];
            if (string.Equals((string)account.SmtpAddress, emailAddress, StringComparison.OrdinalIgnoreCase))
            {
                return account;
            }
        }

        return null;
    }

    private static string BuildBody(TicketRecord record, string signatureHtml, bool includePackageLink)
    {
        var requester = System.Net.WebUtility.HtmlEncode(record.RequesterName);
        var packageName = System.Net.WebUtility.HtmlEncode(record.PackageName);
        var packageContent = includePackageLink
            ? BuildPackageContent(record.PackageLinkUrl, packageName)
            : packageName;

        var signatureBlock = string.IsNullOrWhiteSpace(signatureHtml)
            ? string.Empty
            : $"<br/>{signatureHtml}";

        return $"""
<p>Buen dia.</p>
<p>Se ha dado de alta un nuevo ticket para paso a produccion</p>
<p>Responsable del cambio y validaciones {requester}</p>
<p>Ruta: {packageContent}</p>
{signatureBlock}
""";
    }

    private static string BuildPlainTextBody(TicketRecord record)
    {
        return $"""
Buen dia.

Se ha dado de alta un nuevo ticket para paso a produccion

Responsable del cambio y validaciones {record.RequesterName}
Ruta: {record.PackageName}
""";
    }

    private static string BuildSubject(TicketRecord record)
    {
        var ticket = NormalizeForSubject(record.TicketNumber);
        var package = NormalizeForSubject(record.PackageName);
        var bank = NormalizeForSubject(record.BankName);
        var subject = $"[Nexus Ticket Produccion] {ticket} - {package} - {bank}";

        return subject.Length <= 240 ? subject : subject[..240];
    }

    private static string NormalizeForSubject(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var sanitized = value
            .Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Replace("\t", " ", StringComparison.Ordinal)
            .Trim();

        return string.Join(' ', sanitized.Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    private static string BuildPackageContent(string packageLinkUrl, string packageNameHtml)
    {
        if (!Uri.TryCreate(packageLinkUrl, UriKind.Absolute, out var uri))
        {
            return packageNameHtml;
        }

        if (uri.Scheme != Uri.UriSchemeHttp
            && uri.Scheme != Uri.UriSchemeHttps
            && uri.Scheme != Uri.UriSchemeMailto
            && uri.Scheme != Uri.UriSchemeFile)
        {
            return packageNameHtml;
        }

        return $"<a href=\"{System.Net.WebUtility.HtmlEncode(packageLinkUrl)}\">{packageNameHtml}</a>";
    }

    private static bool IsExpectedRangeError(Exception ex)
    {
        return ex.Message.Contains("Value does not fall within the expected range", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeEmptyBody(string html)
    {
        var normalized = html.Replace("&nbsp;", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("<br>", "", StringComparison.OrdinalIgnoreCase)
            .Replace("<br/>", "", StringComparison.OrdinalIgnoreCase)
            .Replace("<p>", "", StringComparison.OrdinalIgnoreCase)
            .Replace("</p>", "", StringComparison.OrdinalIgnoreCase)
            .Trim();

        return string.IsNullOrWhiteSpace(normalized)
               || normalized.Contains("<body></body>", StringComparison.OrdinalIgnoreCase)
               || normalized.Contains("<body> </body>", StringComparison.OrdinalIgnoreCase);
    }

    private static string ReadSignatureHtmlByName(string signatureName)
    {
        if (string.IsNullOrWhiteSpace(signatureName))
        {
            return string.Empty;
        }

        var signaturesPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Microsoft",
            "Signatures");

        var signatureFilePath = Path.Combine(signaturesPath, $"{signatureName}.htm");
        if (!File.Exists(signatureFilePath))
        {
            return string.Empty;
        }

        return File.ReadAllText(signatureFilePath);
    }

    private static string ExtractSignatureFragment(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return string.Empty;
        }

        var startTagIndex = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
        if (startTagIndex < 0)
        {
            return html.Trim();
        }

        var bodyOpenEndIndex = html.IndexOf('>', startTagIndex);
        if (bodyOpenEndIndex < 0)
        {
            return html.Trim();
        }

        var bodyCloseIndex = html.IndexOf("</body>", bodyOpenEndIndex, StringComparison.OrdinalIgnoreCase);
        if (bodyCloseIndex <= bodyOpenEndIndex)
        {
            return html[(bodyOpenEndIndex + 1)..].Trim();
        }

        return html.Substring(bodyOpenEndIndex + 1, bodyCloseIndex - bodyOpenEndIndex - 1).Trim();
    }

    private static string ReadDefaultSignatureName()
    {
        var officeVersions = new[] { "16.0", "15.0", "14.0" };

        foreach (var version in officeVersions)
        {
            using var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{version}\Common\MailSettings");
            var value = key?.GetValue("NewSignature") as string;
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return string.Empty;
    }

    private static void TrySendWithFallback(dynamic outlookApp, OutlookEmailSettings settings, TicketRecord record, dynamic preparedMail)
    {
        try
        {
            preparedMail.Send();
        }
        catch (Exception ex) when (IsExpectedRangeError(ex))
        {
            if (WasAlreadySent(preparedMail))
            {
                return;
            }

            try
            {
                // Primer fallback de envio: mantenemos firma, simplificando el cuerpo HTML.
                var fallbackWithSignature = PrepareMail(
                    outlookApp,
                    settings,
                    record,
                    includeSignature: true,
                    includePackageLink: false);

                fallbackWithSignature.Send();
            }
            catch (Exception secondEx) when (IsExpectedRangeError(secondEx))
            {
                // Ultimo recurso: correo en texto plano sin firma.
                dynamic fallbackMail = outlookApp.CreateItem(OlMailItem);
                TrySetSenderAccount(outlookApp, fallbackMail, settings.SenderAccount);
                fallbackMail.Subject = BuildSubject(record);
                fallbackMail.Body = BuildPlainTextBody(record);
                AddRecipients(fallbackMail, settings);

                var allResolved = (bool)fallbackMail.Recipients.ResolveAll();
                if (!allResolved)
                {
                    throw new InvalidOperationException("No fue posible resolver uno o mas destinatarios en Outlook.");
                }

                fallbackMail.Send();
            }
        }
    }

    private static bool WasAlreadySent(dynamic mail)
    {
        try
        {
            return (bool)mail.Sent;
        }
        catch
        {
            return false;
        }
    }

}
