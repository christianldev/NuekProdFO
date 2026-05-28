using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using NuekProdFO.Models;
using NuekProdFO.Services;

namespace NuekProdFO;

public partial class MainWindow : Window
{
    private readonly ExcelTicketExtractorService _excelService = new();
    private readonly SmtpEmailService _emailService = new();
    private readonly List<string> _selectedFiles = [];

    public ObservableCollection<TicketRecord> Records { get; } = [];

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
    }

    private void SelectFiles_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            Multiselect = true,
            Title = "Selecciona uno o varios archivos Excel"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        _selectedFiles.Clear();
        _selectedFiles.AddRange(dialog.FileNames);

        TxtFileCount.Text = $"{_selectedFiles.Count} archivos seleccionados";
        SetStatus("Archivos cargados.");
    }

    private async void ProcessFiles_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedFiles.Count == 0)
        {
            MessageBox.Show("Selecciona al menos un archivo Excel.", "Atencion", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var settings = BuildExtractionSettings();

        Records.Clear();
        SetStatus("Procesando archivos...");

        var errors = new List<string>();

        await Task.Run(() =>
        {
            foreach (var file in _selectedFiles)
            {
                try
                {
                    var record = _excelService.ExtractFromFile(file, settings);
                    Dispatcher.Invoke(() => Records.Add(record));
                }
                catch (Exception ex)
                {
                    errors.Add($"{Path.GetFileName(file)}: {ex.Message}");
                }
            }
        });

        if (errors.Count > 0)
        {
            MessageBox.Show(
                "Algunos archivos no se pudieron procesar:\n\n" + string.Join("\n", errors),
                "Advertencia",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        SetStatus($"Procesados: {Records.Count} registros listos.");
    }

    private async void SendEmails_Click(object sender, RoutedEventArgs e)
    {
        if (Records.Count == 0)
        {
            MessageBox.Show("No hay datos para enviar. Primero procesa archivos.", "Atencion", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SmtpSettings smtpSettings;
        try
        {
            smtpSettings = BuildSmtpSettings();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Configuracion invalida", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetStatus("Enviando correos...");
        var sent = 0;
        var failures = new List<string>();

        foreach (var record in Records)
        {
            try
            {
                await _emailService.SendTicketNotificationAsync(smtpSettings, record);
                sent++;
            }
            catch (Exception ex)
            {
                failures.Add($"Ticket {record.TicketNumber} ({record.FileName}): {ex.Message}");
            }
        }

        SetStatus($"Enviados: {sent}/{Records.Count}");

        if (failures.Count > 0)
        {
            MessageBox.Show(
                "Se enviaron algunos correos, pero hubo errores:\n\n" + string.Join("\n", failures),
                "Resultado de envio",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        MessageBox.Show("Todos los correos fueron enviados correctamente.", "Exito", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private ExcelExtractionSettings BuildExtractionSettings()
    {
        return new ExcelExtractionSettings
        {
            SheetName = TxtSheetName.Text,
            TicketCellAddress = TxtTicketCell.Text,
            PackageCellAddress = TxtPackageCell.Text,
            RequesterCellAddress = TxtRequesterCell.Text
        };
    }

    private SmtpSettings BuildSmtpSettings()
    {
        if (string.IsNullOrWhiteSpace(TxtSmtpHost.Text))
        {
            throw new InvalidOperationException("El servidor SMTP es obligatorio.");
        }

        if (!int.TryParse(TxtSmtpPort.Text, out var port) || port <= 0)
        {
            throw new InvalidOperationException("El puerto SMTP no es valido.");
        }

        var toRecipients = ParseEmailList(TxtTo.Text);
        var ccRecipients = ParseEmailList(TxtCc.Text);

        if (string.IsNullOrWhiteSpace(TxtFrom.Text))
        {
            throw new InvalidOperationException("El correo remitente es obligatorio.");
        }

        if (toRecipients.Count == 0)
        {
            throw new InvalidOperationException("Debes ingresar al menos un destinatario.");
        }

        return new SmtpSettings
        {
            Host = TxtSmtpHost.Text.Trim(),
            Port = port,
            UseSsl = ChkSsl.IsChecked == true,
            Username = TxtSmtpUser.Text.Trim(),
            Password = TxtSmtpPassword.Password,
            From = TxtFrom.Text.Trim(),
            ToRecipients = toRecipients,
            CcRecipients = ccRecipients
        };
    }

    private static List<string> ParseEmailList(string rawList)
    {
        return rawList
            .Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private void SetStatus(string message)
    {
        TxtStatus.Text = message;
    }
}