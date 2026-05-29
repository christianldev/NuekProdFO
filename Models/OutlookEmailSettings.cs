namespace NuekProdFO.Models;

public class OutlookEmailSettings
{
    public string SenderAccount { get; set; } = string.Empty;

    public List<string> ToRecipients { get; set; } = [];

    public List<string> CcRecipients { get; set; } = [];

    public bool HideDraftWhenApplyingSignature { get; set; } = false;
}
