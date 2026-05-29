using System.IO;
using NuekProdFO.Models;

namespace NuekProdFO.Services;

public class EnvironmentConfigService
{
    public OutlookEmailSettings LoadOutlookSettings(string envPath)
    {
        if (!File.Exists(envPath))
        {
            throw new InvalidOperationException($"No se encontro el archivo .env en: {envPath}");
        }

        var variables = ParseEnvFile(envPath);

        var sender = Require(variables, "OUTLOOK_SENDER");
        var toRaw = Require(variables, "OUTLOOK_TO");
        var hideDraftRaw = ReadOptional(variables, "OUTLOOK_HIDE_DRAFT", "false");
        var hideDraft = bool.TryParse(hideDraftRaw, out var parsedHideDraft) && parsedHideDraft;

        return new OutlookEmailSettings
        {
            SenderAccount = sender,
            ToRecipients = ParseEmailList(toRaw),
            CcRecipients = ParseEmailList(ReadOptional(variables, "OUTLOOK_CC", string.Empty)),
            HideDraftWhenApplyingSignature = hideDraft
        };
    }

    public SmtpSettings LoadSmtpSettings(string envPath)
    {
        if (!File.Exists(envPath))
        {
            throw new InvalidOperationException($"No se encontro el archivo .env en: {envPath}");
        }

        var variables = ParseEnvFile(envPath);

        var host = Require(variables, "SMTP_HOST");
        var from = Require(variables, "SMTP_FROM");
        var toRaw = Require(variables, "SMTP_TO");
        var username = Require(variables, "SMTP_USER");
        var password = Require(variables, "SMTP_PASS");

        var portRaw = ReadOptional(variables, "SMTP_PORT", "587");
        if (!int.TryParse(portRaw, out var port) || port <= 0)
        {
            throw new InvalidOperationException("SMTP_PORT no es valido en .env");
        }

        var sslRaw = ReadOptional(variables, "SMTP_SSL", "true");
        var useSsl = bool.TryParse(sslRaw, out var parsedSsl) && parsedSsl;

        return new SmtpSettings
        {
            Host = host,
            Port = port,
            UseSsl = useSsl,
            Username = username,
            Password = password,
            From = from,
            ToRecipients = ParseEmailList(toRaw),
            CcRecipients = ParseEmailList(ReadOptional(variables, "SMTP_CC", string.Empty))
        };
    }

    private static Dictionary<string, string> ParseEnvFile(string envPath)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var rawLine in File.ReadAllLines(envPath))
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith('#'))
            {
                continue;
            }

            var separator = line.IndexOf('=');
            if (separator <= 0)
            {
                continue;
            }

            var key = line[..separator].Trim();
            var value = line[(separator + 1)..].Trim();

            if (value.StartsWith('"') && value.EndsWith('"') && value.Length >= 2)
            {
                value = value[1..^1];
            }

            map[key] = value;
        }

        return map;
    }

    private static string Require(Dictionary<string, string> variables, string key)
    {
        if (!variables.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            throw new InvalidOperationException($"Falta la variable obligatoria {key} en .env");
        }

        return value.Trim();
    }

    private static string ReadOptional(Dictionary<string, string> variables, string key, string defaultValue)
    {
        if (!variables.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            return defaultValue;
        }

        return value.Trim();
    }

    private static List<string> ParseEmailList(string rawList)
    {
        return rawList
            .Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
