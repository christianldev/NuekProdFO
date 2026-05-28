namespace NuekProdFO.Models;

public class SmtpSettings
{
    public string Host { get; set; } = string.Empty;

    public int Port { get; set; } = 587;

    public bool UseSsl { get; set; } = true;

    public string Username { get; set; } = string.Empty;

    public string Password { get; set; } = string.Empty;

    public string From { get; set; } = string.Empty;

    public List<string> ToRecipients { get; set; } = [];

    public List<string> CcRecipients { get; set; } = [];
}
