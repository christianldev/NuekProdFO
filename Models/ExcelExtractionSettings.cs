namespace NuekProdFO.Models;

public class ExcelExtractionSettings
{
    public string? SheetName { get; set; }

    public string TicketCellAddress { get; set; } = "D8";

    public string PackageCellAddress { get; set; } = "B59";

    public string RequesterCellAddress { get; set; } = "D10";

    public string BankCellAddress { get; set; } = "H10";
}
