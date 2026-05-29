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

    public void SendTicketNotification(OutlookEmailSettings settings, TicketRecord record)
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application", throwOnError: true)!;
        dynamic outlookApp = Activator.CreateInstance(outlookType)!;

        dynamic mail = outlookApp.CreateItem(OlMailItem);

        mail.Subject = $"[Nexus Ticket Producción] {record.TicketNumber} - {record.PackageName} - {record.BankName}";
        mail.HTMLBody = BuildBody(record);

        foreach (var address in settings.ToRecipients)
        {
            mail.Recipients.Add(address);
        }

        foreach (var address in settings.CcRecipients)
        {
            dynamic recipient = mail.Recipients.Add(address);
            recipient.Type = OlCC;
        }

        mail.Recipients.ResolveAll();

        if (!string.IsNullOrWhiteSpace(settings.SenderAccount))
        {
            dynamic? account = FindAccount(outlookApp, settings.SenderAccount);
            if (account is not null)
            {
                mail.SendUsingAccount = account;
            }
        }

        mail.Send();
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

    private static string BuildBody(TicketRecord record)
    {
        var requester = System.Net.WebUtility.HtmlEncode(record.RequesterName);
        var packageName = System.Net.WebUtility.HtmlEncode(record.PackageName);
        var packageContent = string.IsNullOrWhiteSpace(record.PackageLinkUrl)
            ? packageName
            : $"<a href=\"{System.Net.WebUtility.HtmlEncode(record.PackageLinkUrl)}\">{packageName}</a>";

        return $"""
<p>Buen dia.</p>
<p>Se ha dado de alta un nuevo ticket para paso a produccion</p>
<p>Responsable del cambio y validaciones {requester}</p>
<p>Ruta: {packageContent}</p>
""";
    }
}
