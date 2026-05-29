using ClosedXML.Excel;
using NuekProdFO.Models;

namespace NuekProdFO.Services;

public class ExcelTicketExtractorService
{
    public TicketRecord ExtractFromFile(string filePath, ExcelExtractionSettings settings)
    {
        using var workbook = new XLWorkbook(filePath);

        var worksheet = ResolveWorksheet(workbook, settings);
        if (worksheet is null)
        {
            throw new InvalidOperationException("No se encontro la hoja especificada en el archivo.");
        }

        return new TicketRecord
        {
            FilePath = filePath,
            TicketNumber = ReadCell(worksheet, settings.TicketCellAddress),
            PackageName = ReadCell(worksheet, settings.PackageCellAddress),
            PackageLinkUrl = ReadHyperlink(worksheet, settings.PackageLinkCellAddress),
            RequesterName = ReadCell(worksheet, settings.RequesterCellAddress),
            BankName = ReadCell(worksheet, settings.BankCellAddress)
        };
    }

    private static IXLWorksheet? ResolveWorksheet(XLWorkbook workbook, ExcelExtractionSettings settings)
    {
        if (!string.IsNullOrWhiteSpace(settings.SheetName))
        {
            return workbook.Worksheets.FirstOrDefault(x =>
                string.Equals(x.Name, settings.SheetName, StringComparison.OrdinalIgnoreCase));
        }

        // Si no se indica hoja, buscamos la que mejor coincide con las celdas configuradas.
        // Esto reduce errores cuando la primera hoja no es la de datos.
        var targetAddresses = new[]
        {
            settings.TicketCellAddress,
            settings.PackageCellAddress,
            settings.RequesterCellAddress,
            settings.BankCellAddress
        };

        var bestMatch = workbook.Worksheets
            .Where(x => x.Visibility == XLWorksheetVisibility.Visible)
            .Select(x => new
            {
                Sheet = x,
                Score = targetAddresses.Count(address => !string.IsNullOrWhiteSpace(ReadCell(x, address)))
            })
            .OrderByDescending(x => x.Score)
            .FirstOrDefault();

        if (bestMatch is not null && bestMatch.Score > 0)
        {
            return bestMatch.Sheet;
        }

        return workbook.Worksheets.FirstOrDefault(x => x.Visibility == XLWorksheetVisibility.Visible)
               ?? workbook.Worksheets.FirstOrDefault();
    }

    private static string ReadCell(IXLWorksheet worksheet, string cellAddress)
    {
        var normalizedAddress = string.IsNullOrWhiteSpace(cellAddress) ? "A1" : cellAddress.Trim();
        var cell = worksheet.Cell(normalizedAddress);

        if (cell.IsMerged())
        {
            cell = cell.MergedRange().FirstCell();
        }

        // GetFormattedString suele representar mejor el valor mostrado en Excel
        // cuando hay formatos o formulas.
        return cell.GetFormattedString().Trim();
    }

    private static string ReadHyperlink(IXLWorksheet worksheet, string cellAddress)
    {
        var normalizedAddress = string.IsNullOrWhiteSpace(cellAddress) ? "I59" : cellAddress.Trim();
        var cell = worksheet.Cell(normalizedAddress);

        if (cell.IsMerged())
        {
            cell = cell.MergedRange().FirstCell();
        }

        if (cell.HasHyperlink)
        {
            var hyperlink = cell.GetHyperlink();

            if (hyperlink.ExternalAddress is not null)
            {
                return hyperlink.ExternalAddress.ToString();
            }

            if (!string.IsNullOrWhiteSpace(hyperlink.InternalAddress))
            {
                return hyperlink.InternalAddress.Trim();
            }
        }

        // Fallback: si la celda tiene texto y ese texto ya es URL absoluta.
        var rawValue = cell.GetFormattedString().Trim();
        if (Uri.TryCreate(rawValue, UriKind.Absolute, out _))
        {
            return rawValue;
        }

        return string.Empty;
    }
}
