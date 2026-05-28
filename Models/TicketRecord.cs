using System.IO;

namespace NuekProdFO.Models;

public class TicketRecord
{
    public string FilePath { get; set; } = string.Empty;

    public string FileName => Path.GetFileName(FilePath);

    public string TicketNumber { get; set; } = string.Empty;

    public string PackageName { get; set; } = string.Empty;

    public string RequesterName { get; set; } = string.Empty;

    public string BankName { get; set; } = string.Empty;
}
