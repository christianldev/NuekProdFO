using System.Net;
using System.Net.Mail;
using NuekProdFO.Models;

namespace NuekProdFO.Services;

public class SmtpEmailService
{
    public async Task SendTicketNotificationAsync(SmtpSettings settings, TicketRecord record)
    {
        ValidateProviderRequirements(settings);

        using var client = new SmtpClient(settings.Host, settings.Port)
        {
            EnableSsl = settings.UseSsl,
            UseDefaultCredentials = false,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            Timeout = 30000,
            Credentials = new NetworkCredential(settings.Username, settings.Password)
        };

        var fromAddress = string.IsNullOrWhiteSpace(settings.From)
            ? settings.Username
            : settings.From;

        using var message = new MailMessage
        {
            From = new MailAddress(fromAddress),
            Subject = $"Ticket {record.TicketNumber} | {record.PackageName}",
                    Body = BuildBody(record),
                    IsBodyHtml = true
        };

        foreach (var address in settings.ToRecipients)
        {
            message.To.Add(address);
        }

        foreach (var address in settings.CcRecipients)
        {
            message.CC.Add(address);
        }

        try
        {
            await client.SendMailAsync(message);
        }
        catch (SmtpException ex) when (IsAuthError(ex))
        {
            throw new InvalidOperationException(
                "Autenticacion SMTP rechazada (535/5.7.x). Verifica en .env: SMTP_HOST=smtp.office365.com, SMTP_PORT=587, SMTP_SSL=true, SMTP_USER y SMTP_PASS correctos. " +
                "En Microsoft 365 habilita SMTP AUTH para el buzon y usa una contrasena de aplicacion si tienes MFA. " +
                "Ademas, SMTP_FROM debe ser el mismo correo autenticado o un remitente con permiso Send As.",
                ex);
        }
    }

    private static void ValidateProviderRequirements(SmtpSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.Host))
        {
            throw new InvalidOperationException("SMTP_HOST es obligatorio.");
        }

        var isOffice365 = settings.Host.Contains("office365", StringComparison.OrdinalIgnoreCase)
                         || settings.Host.Contains("outlook", StringComparison.OrdinalIgnoreCase);

        if (!isOffice365)
        {
            return;
        }

        if (settings.Port != 587)
        {
            throw new InvalidOperationException("Para Microsoft 365 el puerto SMTP debe ser 587.");
        }

        if (!settings.UseSsl)
        {
            throw new InvalidOperationException("Para Microsoft 365 debes usar SMTP_SSL=true (STARTTLS).\n");
        }
    }

    private static bool IsAuthError(SmtpException ex)
    {
        return ex.Message.Contains("5.7.57", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("535", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("authentication", StringComparison.OrdinalIgnoreCase)
               || ex.Message.Contains("not authenticated", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildBody(TicketRecord record)
    {
        var ticketNumber = WebUtility.HtmlEncode(record.TicketNumber);
        var requester = WebUtility.HtmlEncode(record.RequesterName);
        var fileName = WebUtility.HtmlEncode(record.FileName);
        var packageName = WebUtility.HtmlEncode(record.PackageName);

        var packageContent = string.IsNullOrWhiteSpace(record.PackageLinkUrl)
            ? packageName
            : $"<a href=\"{WebUtility.HtmlEncode(record.PackageLinkUrl)}\">{packageName}</a>";

        return $"""
<p>Se notifica el siguiente ticket procesado desde Excel:</p>
<p>Ticket: {ticketNumber}<br/>
Paquete: {packageContent}<br/>
Responsable de solicitud: {requester}<br/>
Archivo fuente: {fileName}</p>
<p>Mensaje generado automaticamente por NuekProdFO.</p>
""";
    }
}
