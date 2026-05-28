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
    private readonly OutlookEmailService _emailService = new();
    private readonly EnvironmentConfigService _environmentConfigService = new();
    private readonly List<string> _selectedFiles = [];

    public ObservableCollection<TicketRecord> Records { get; } = [];

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        LoadOutlookConfigIntoUi();
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

        OutlookEmailSettings outlookSettings;
        try
        {
            outlookSettings = BuildOutlookSettingsFromEnv();
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
                await Task.Run(() => _emailService.SendTicketNotification(outlookSettings, record));
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

    private OutlookEmailSettings BuildOutlookSettingsFromEnv()
    {
        var envPath = ResolveEnvPath();

        var settings = _environmentConfigService.LoadOutlookSettings(envPath);
        if (settings.ToRecipients.Count == 0)
        {
            throw new InvalidOperationException("OUTLOOK_TO no tiene destinatarios validos en .env");
        }

        return settings;
    }

    private void LoadOutlookConfigIntoUi()
    {
        try
        {
            var envPath = ResolveEnvPath();
            var settings = BuildOutlookSettingsFromEnv();

            TxtSmtpHost.Text = "Outlook local";
            TxtSmtpPort.Text = "(COM)";
            ChkSsl.IsChecked = true;
            TxtSmtpUser.Text = settings.SenderAccount;
            TxtSmtpPassword.Password = string.Empty;
            TxtFrom.Text = settings.SenderAccount;
            TxtTo.Text = string.Join(";", settings.ToRecipients);
            TxtCc.Text = string.Join(";", settings.CcRecipients);

            SetStatus($"Outlook COM listo · {settings.SenderAccount}");
        }
        catch
        {
            SetStatus("Configura .env: OUTLOOK_SENDER y OUTLOOK_TO");
        }
    }

    private static string ResolveEnvPath()
    {
        var explicitPath = Environment.GetEnvironmentVariable("NUEKPRODFO_ENV_PATH");
        if (!string.IsNullOrWhiteSpace(explicitPath) && File.Exists(explicitPath))
        {
            return explicitPath;
        }

        var candidatePaths = new[]
        {
            Path.Combine(Environment.CurrentDirectory, ".env"),
            Path.Combine(AppContext.BaseDirectory, ".env")
        };

        var found = candidatePaths.FirstOrDefault(File.Exists);
        if (found is not null)
        {
            return found;
        }

        throw new InvalidOperationException(
            "No se encontro .env. Se busco en el directorio actual y en el directorio de ejecucion.");
    }

    private void SetStatus(string message)
    {
        TxtStatus.Text = message;
    }
}