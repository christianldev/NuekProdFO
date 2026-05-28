using System.Net;
using System.Net.Mail;
using NuekProdFO.Models;

namespace NuekProdFO.Services;

public class SmtpEmailService
{
    public async Task SendTicketNotificationAsync(SmtpSettings settings, TicketRecord record)
    {
        using var client = new SmtpClient(settings.Host, settings.Port)
        {
            EnableSsl = settings.UseSsl,
            Credentials = new NetworkCredential(settings.Username, settings.Password)
        };

        using var message = new MailMessage
        {
            From = new MailAddress(settings.From),
            Subject = $"Ticket {record.TicketNumber} | {record.PackageName}",
            Body = BuildBody(record),
            IsBodyHtml = false
        };

        foreach (var address in settings.ToRecipients)
        {
            message.To.Add(address);
        }

        foreach (var address in settings.CcRecipients)
        {
            message.CC.Add(address);
        }

        await client.SendMailAsync(message);
    }

    private static string BuildBody(TicketRecord record)
    {
        return $"""
Se notifica el siguiente ticket procesado desde Excel:

Ticket: {record.TicketNumber}
Paquete: {record.PackageName}
Responsable de solicitud: {record.RequesterName}
Archivo fuente: {record.FileName}

Mensaje generado automaticamente por NuekProdFO.
""";
    }
}
