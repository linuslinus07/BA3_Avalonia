using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using BA3.Avalonia.Settings;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;

// 04.04.2026

namespace BA3.Avalonia;

public partial class MainWindow : Window
{
    private readonly SettingsService _settingsService = new(appName: "BA3Code");
    private string? _folderpath;

    public string version { get; set; } = "v010426";

    public MainWindow()
    {
        InitializeComponent();

        DataContext = this;

        Opened += async (_, _) => await LoadSettingsAsync();
        Closing += OnWindowClosing;
    }

    private bool _isClosing;
    private async void OnWindowClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_isClosing) return;

        e.Cancel = true;
        _isClosing = true;

        // Capture all UI values on the UI thread FIRST
        var pathInput = _pathInput.Text ?? "";
        System.Diagnostics.Debug.WriteLine($"OnWindowClosing - pathInput: '{pathInput}'");
        
        // Use double.TryParse for safety
        double.TryParse(_height.Text, out var height);
        double.TryParse(_mat.Text, out var mat);
        double.TryParse(_diaInt.Text, out var diaInt);
        double.TryParse(_xErste.Text, out var xErste);
        double.TryParse(_rabo.Text, out var rabo);
        double.TryParse(_lab.Text, out var lab);
        double.TryParse(_deg.Text, out var deg);
        double.TryParse(_xVersch.Text, out var xVersch);
        double.TryParse(_secur.Text, out var secur);
        double.TryParse(_durch.Text, out var durch);
        var spinRpm = _spinRPM.Value;
        double.TryParse(_xAbst.Text, out var xAbst);
        var custom = _custom.IsChecked == true;
        double.TryParse(_anz.Text, out var anz);
        var farbe = _farbe.IsChecked == true;
        var versetzt = _versetzt.IsChecked == true;
        var kunde = _firmaName.Text ?? "";
        var keepValues = _speichern.IsChecked == true;
        
        var selectedDrill = OptionA.IsChecked == true ? "0.8" :
                        OptionB.IsChecked == true ? "1.0" :
                        OptionC.IsChecked == true ? "1.3" :
                        OptionD.IsChecked == true ? "1.5" :
                        OptionE.IsChecked == true ? "2.0" :
                        OptionF.IsChecked == true ? "3.0" :
                        "None";
        
        try
        {
            var s = await _settingsService.LoadAsync();
            s.PathInput = pathInput;
            
            if (keepValues)
            {
                s.Height = height;
                s.Mat = mat;
                s.DiaInt = diaInt;
                s.XErste = xErste;
                s.Rabo = rabo;
                s.Lab = lab;
                s.XVersch = xVersch;
                s.Secur = secur;
                s.Durch = durch;
                s.SpinRpm = spinRpm;
                s.Deg = deg;
                s.XAbst = xAbst;
                s.Custom = custom;
                s.Anz = anz;
                s.Farbe = farbe;
                s.Versetzt = versetzt;
                s.Kunde = kunde;
                s.DrillDia = selectedDrill;
                s.KeepValues = true;
            }
            else
            {
                s.KeepValues = false;
            }
            await _settingsService.SaveAsync(s);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"ERROR SAVING: {ex}");
        }
        finally
        {
            Close();
        }
    }
    
    private async Task LoadSettingsAsync()
    {
        var s = await _settingsService.LoadAsync();
        
        //await Task.Delay(500);  // ✅ Increase delay significantly

        System.Diagnostics.Debug.WriteLine($"Full loaded settings: {System.Text.Json.JsonSerializer.Serialize(s)}");
        System.Diagnostics.Debug.WriteLine($"Loaded PathInput raw: '{s.PathInput}'");
        System.Diagnostics.Debug.WriteLine($"Loaded PathInput is null: {s.PathInput == null}");
        System.Diagnostics.Debug.WriteLine($"Loaded PathInput length: {s.PathInput?.Length}");

        _pathInput.Text = s.PathInput ?? "";

        if (!s.KeepValues)
            return;

        _height.Text = s.Height.ToString();
        _mat.Text = s.Mat.ToString();
        _diaInt.Text = s.DiaInt.ToString();
        _xErste.Text = s.XErste.ToString();
        _rabo.Text = s.Rabo.ToString();
        _lab.Text = s.Lab.ToString();
        _xVersch.Text = s.XVersch.ToString();
        _secur.Text = s.Secur.ToString();
        _durch.Text = s.Durch.ToString();
        _spinRPM.Value = s.SpinRpm;
        _deg.Text = s.Deg.ToString();
        _xAbst.Text = s.XAbst.ToString();
        _custom.IsChecked = s.Custom;
        _anz.Text = s.Anz.ToString();
        _speichern.IsChecked = true;
        _farbe.IsChecked = s.Farbe;
        _versetzt.IsChecked = s.Versetzt;
        _firmaName.Text = s.Kunde ?? "";

        OptionA.IsChecked = s.DrillDia == "0.8";
        OptionB.IsChecked = s.DrillDia == "1.0";
        OptionC.IsChecked = s.DrillDia == "1.3";
        OptionD.IsChecked = s.DrillDia == "1.5";
        OptionE.IsChecked = s.DrillDia == "2.0";
        OptionF.IsChecked = s.DrillDia == "3.0";
    }

    /*private async Task PersistPathAsync()
    {
        var s = await _settingsService.LoadAsync();
        s.PathInput = _pathInput.Text;
        await _settingsService.SaveAsync(s);
    }*/

    private void ClearInputs(object? sender, RoutedEventArgs e)
    {
        _diaInt.Text = "0";
        _height.Text = "0";
        _lab.Text = "4";
        _xErste.Text = "0";
        _deg.Text = "1";
        _mat.Text = "5";
        _xAbst.Text = "0";
        _rabo.Text = "0";
        _xVersch.Text = "0";
        _anz.Text = "1";
        _secur.Text = "8";
        _durch.Text = "2";
        _firmaName.Text = "";
        _spinRPM.Value = 24000;
        _custom.IsChecked = false;
        _farbe.IsChecked = false;
        _versetzt.IsChecked = true;
        _speichern.IsChecked = false;

        OptionA.IsChecked = false;
        OptionB.IsChecked = false;
        OptionC.IsChecked = false;
        OptionD.IsChecked = false;
        OptionE.IsChecked = false;
        OptionF.IsChecked = false;
    }

    private async void browse(object? sender, RoutedEventArgs e)
    {
        var options = new FolderPickerOpenOptions
        {
            Title = "Ordner auswählen",
            AllowMultiple = false
        };

        var folders = await StorageProvider.OpenFolderPickerAsync(options);
        if (folders.Count <= 0)
            return;

        _folderpath = folders[0].Path.LocalPath;
        _pathInput.Text = _folderpath;
    }

    private static Task ShowInfoAsync(Window parent, string title, string message)
    {
        return MessageBoxManager
            .GetMessageBoxStandard(title, message, ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Info, WindowStartupLocation.CenterOwner)
            .ShowWindowDialogAsync(parent);
    }

    private static Task ShowErrorAsync(Window parent, string title, string message)
    {
        return MessageBoxManager
            .GetMessageBoxStandard(title, message, ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error, WindowStartupLocation.CenterOwner)
            .ShowWindowDialogAsync(parent);
    }

    private static async Task<bool> ConfirmOkCancelAsync(Window parent, string title, string message)
    {
        var result = await MessageBoxManager
            .GetMessageBoxStandard(title, message, ButtonEnum.OkCancel, MsBox.Avalonia.Enums.Icon.Warning, WindowStartupLocation.CenterOwner)
            .ShowWindowDialogAsync(parent);

        return result == ButtonResult.Ok;
    }

    public async void Calculate(object? sender, RoutedEventArgs e)
    {
        #region -- Calculate --
        try
        {
            #region -- Variabeln --
            if (!double.TryParse(_diaInt.Text, out double diaInt))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für den Innendurchmesser ein.");
                return;
            }

            if (!double.TryParse(_height.Text, out double height))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für die Höhe ein.");
                return;
            }

            if (!double.TryParse(_lab.Text, out double lab))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für den Lochabstand ein.");
                return;
            }

            if (!double.TryParse(_xErste.Text, out double xErste))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für den Randabstand ein.");
                return;
            }

            if (!double.TryParse(_deg.Text, out double deg))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für den Lochabstand Radial ein.");
                return;
            }

            if (!double.TryParse(_mat.Text, out double mat))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für die Materialstärke ein.");
                return;
            }

            if (!double.TryParse(_xAbst.Text, out double xAbst))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für den X-Abstand ein.");
                return;
            }

            if (!double.TryParse(_rabo.Text, out double rabo))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für den Randabstand ein.");
                return;
            }

            if (!double.TryParse(_xVersch.Text, out double xVersch))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für die X-Verschiebung ein.");
                return;
            }

            if (!double.TryParse(_anz.Text, out double anz))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für die Anzahl ein.");
                return;
            }

            if (!double.TryParse(_secur.Text, out double secur))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für die Sicherheitshöhe ein.");
                return;
            }

            if (!double.TryParse(_durch.Text, out double durch))
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Bitte gib eine gültige Zahl für die Durchbruchtiefe ein.");
                return;
            }

            bool custom = _custom.IsChecked == true;
            double spinRPM = _spinRPM.Value;
            bool IsChecked = _speichern.IsChecked == true;
            bool farbe = _farbe.IsChecked == true;
            bool offsetbtn = _versetzt.IsChecked == true;

            string firma = _firmaName.Text ?? "";
            string kunde = "_" + firma;

            string directory = _pathInput.Text ?? "";
            string fileName = fileName = $"{anz}x_d={diaInt}mm_h={height}mm_t={mat}{kunde}.nc";

            if ((_firmaName.Text ?? "") == "")
            {
                fileName = $"{anz}x_d={diaInt}mm_h={height}mm_t={mat}.nc";
            }

            string filePath = Path.Combine(directory, fileName);

            double drillDia = OptionA.IsChecked == true ? 0.8 :
                      OptionB.IsChecked == true ? 1.0 :
                      OptionC.IsChecked == true ? 1.3 :
                      OptionD.IsChecked == true ? 1.5 :
                      OptionE.IsChecked == true ? 2.0 :
                      OptionF.IsChecked == true ? 3.0 : 0.0;

            if (custom == true)
            {
                spinRPM = _spinRPM.Value;
            }
            else if (drillDia == 0.8)
            {
                if (farbe == true)
                {
                    spinRPM = 33750;
                }
                else
                {
                    spinRPM = 37500;
                }
            }
            else if (drillDia == 1.0)
            {
                if (farbe == true)
                {
                    spinRPM = 27000;
                }
                else
                {
                    spinRPM = 30000;
                }
            }
            else if (drillDia == 1.3)
            {
                if (farbe == true)
                {
                    spinRPM = 21600;
                }
                else
                {
                    spinRPM = 24000;
                }
            }
            else if (drillDia == 1.5)
            {
                if (farbe == true)
                {
                    spinRPM = 18000;
                }
                else
                {
                    spinRPM = 20000;
                }
            }
            else if (drillDia == 2.0)
            {
                spinRPM = 20000;
            }
            else if (drillDia == 3.0)
            {
                spinRPM = 20000;
            }

            double Currentx;
            double xPos;
            double aPos;
            int h;
            int i;
            bool Spind1;
            bool Spind2;
            bool Spind3;
            bool Spind4;

            double diaExt = diaInt + (2 * mat);
            double radInt = diaInt / 2;
            double radExt = diaExt / 2;
            double zStart = radExt + secur;
            double zEnd = diaInt / 2 - durch;

            double nettoBohr = height - rabo - xErste;

            double labOffset = lab / 2;
            double aux1 = nettoBohr / labOffset;
            double floor = Math.Floor(aux1);
            double nbrRowsTot = floor + 1;
            double aux2 = nbrRowsTot / 2;
            double nbrRows = Math.Ceiling(aux2);
            double nbrRowsOffset = Math.Floor(aux2);

            double nbrHolesRad = 360 / deg;
            double aNull = deg - deg;

            double RPD = 60 / lab;
            double Netto = ((nettoBohr / lab) + 1) / RPD;
            double LastL = Math.Round((Netto - Math.Floor(Netto)) * RPD);
            double LetzteL = ((Netto * 15) * lab);

            double anzOff = height + xAbst;


            #endregion

            #region -- Warnungen --
            if (lab == 0)
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Reihen-Abstand darf nicht 0 sein");
                return;
            }

            if (deg == 0)
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Lochabstand zu klein");
                return;
            }

            if (mat == 0)
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Materialstärke darf nicht 0 sein");
                return;
            }

            if (anz == 0)
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Anzahl darf nicht 0 sein");
                return;
            }

            if (secur == 0)
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Sicherheitshöhe darf nicht 0 sein");
                return;
            }

            if (durch == 0)
            {
                var ok = await ConfirmOkCancelAsync(this, "Achtung!", "Duchbruchtiefe bei 0");

                if (!ok)
                {
                    return;
                }
            }

            if (spinRPM == 0)
            {
                var ok = await ConfirmOkCancelAsync(this, "Achtung!", "Drehzahl bei 0");

                if (!ok)
                {
                    return;
                }
            }

            if (drillDia == 0)
            {
                await ShowErrorAsync(this, "Ungültige Eingabe", "Kein Bohrdurchmesser ausgewählt");
                return;
            }
            #endregion

            await Task.Run(() =>
            {
                #region -- Logic --

                using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.ASCII))
                {
                    writer.WriteLine("(" + fileName + ")");
                    writer.WriteLine("(d=" + diaInt + " | h=" + height + " | netto=" + nettoBohr + " | t=" + mat + " | °rad=" + deg + " | #rad=" + nbrHolesRad + " | Lochabstand=" + lab + " | #rows=" + nbrRows + "/" + nbrRowsOffset + " | spin=" + spinRPM + ")");
                    writer.WriteLine("(Verschiebung xNull: " + xVersch + ")");
                    writer.WriteLine("(Anzahl Formen: " + anz + " " + "Abstand: " + xAbst + ")");
                    writer.WriteLine("(T1 D=1. CR=0. TAPER=140DEG-Bohrer)");
                    writer.WriteLine("(Letzte Reihe bei: " + Math.Floor(LastL) + ")");
                    writer.WriteLine("(Netto " + Netto + "/" + Math.Floor(Netto) + ")");
                    writer.WriteLine("(Nettobohr " + nettoBohr + ")");
                    writer.WriteLine("(rpd " + RPD + ")");
                    writer.WriteLine("G90 G94 G91.1 G40 G49 G17");
                    writer.WriteLine("G21");
                    writer.WriteLine("G28 G91 Z0.");
                    writer.WriteLine("G90");

                    writer.WriteLine("m110");                           //alle hoch
                    writer.WriteLine("m101");                           //alle runter (one by one)
                    writer.WriteLine("m102");
                    writer.WriteLine("m103");
                    writer.WriteLine("m104");

                    //writer.WriteLine("M5");                           //(=Spindelstopp, braucht es nicht)
                    writer.WriteLine("T1 M6");                          //Werkzeug 1, M6=Wz-Wechsel
                    writer.WriteLine("S" + spinRPM + " M3");            //Spindeldrehzahl
                    writer.WriteLine("G54");                            //Werkstücknullpunkt"

                    Spind1 = true;
                    Spind2 = true;
                    Spind3 = true;
                    Spind4 = true;
                    aPos = aNull;
                    Currentx = 0;

                    writer.WriteLine("G0 A0");
                    writer.WriteLine("G0 X23");

                    for (int j = 1; j <= anz; j++) //loop for multiple forms
                    {
                        double offset = (j - 1) * anzOff;

                        /*if (j > 1)                                              //wenn mehr als eine form, dann alle spindeln hoch und wieder runter (one by one) damit die nächste form gebohrt werden kann
                        {
                            if (!Spind1) { writer.WriteLine("m101"); Spind1 = true; }
                            if (!Spind2) { writer.WriteLine("m102"); Spind2 = true; }
                            if (!Spind3) { writer.WriteLine("m103"); Spind3 = true; }
                            if (!Spind4) { writer.WriteLine("m104"); Spind4 = true; }
                        }*/

                        if (Netto <= 4)
                        {
                            if (Netto == Math.Floor(Netto))                 //wenn netto eine gerade zahl ist
                            {
                                if (Netto == 1)                             //richtige spindeln auswählen
                                {
                                    if (!Spind1)
                                    {
                                        writer.WriteLine("m101");
                                        Spind1 = true;
                                    }

                                    if (Spind2)
                                    {
                                        writer.WriteLine("m102");
                                        Spind2 = false;
                                    }

                                    if (Spind3)
                                    {
                                        writer.WriteLine("m103");
                                        Spind3 = false;
                                    }

                                    if (Spind4)
                                    {
                                        writer.WriteLine("m104");
                                        Spind4 = false;
                                    }
                                }
                                else if (Netto == 2)
                                {
                                    if (!Spind1)
                                    {
                                        writer.WriteLine("m101");
                                        Spind1 = true;
                                    }

                                    if (!Spind2)
                                    {
                                        writer.WriteLine("m102");
                                        Spind2 = true;
                                    }

                                    if (Spind3)
                                    {
                                        writer.WriteLine("m103");
                                        Spind3 = false;
                                    }

                                    if (Spind4)
                                    {
                                        writer.WriteLine("m104");
                                        Spind4 = false;
                                    }
                                }
                                else if (Netto == 3)
                                {
                                    if (!Spind1)
                                    {
                                        writer.WriteLine("m101");
                                        Spind1 = true;
                                    }

                                    if (!Spind2)
                                    {
                                        writer.WriteLine("m102");
                                        Spind2 = true;
                                    }

                                    if (!Spind3)
                                    {
                                        writer.WriteLine("m103");
                                        Spind3 = true;
                                    }

                                    if (Spind4)
                                    {
                                        writer.WriteLine("m104");
                                        Spind4 = false;
                                    }
                                }
                            }
                            if (Math.Floor(Netto) == 0 && Netto != 0)                 //wenn netto ungerade ist, dann die richtigen spindeln auswählen
                            {
                                if (!Spind1)
                                {
                                    writer.WriteLine("m101");
                                    Spind1 = true;
                                }

                                if (Spind2)
                                {
                                    writer.WriteLine("m102");
                                    Spind2 = false;
                                }

                                if (Spind3)
                                {
                                    writer.WriteLine("m103");
                                    Spind3 = false;
                                }

                                if (Spind4)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                            }
                            else if (Math.Floor(Netto) == 1 && Netto != 1)
                            {
                                if (!Spind1)
                                {
                                    writer.WriteLine("m101");
                                    Spind1 = true;
                                }

                                if (!Spind2)
                                {
                                    writer.WriteLine("m102");
                                    Spind2 = true;
                                }

                                if (Spind3)
                                {
                                    writer.WriteLine("m103");
                                    Spind3 = false;
                                }

                                if (Spind4)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                            }
                            else if (Math.Floor(Netto) == 2 && Netto != 2)
                            {
                                if (!Spind1)
                                {
                                    writer.WriteLine("m101");
                                    Spind1 = true;
                                }

                                if (!Spind2)
                                {
                                    writer.WriteLine("m102");
                                    Spind2 = true;
                                }

                                if (!Spind3)
                                {
                                    writer.WriteLine("m103");
                                    Spind3 = true;
                                }

                                if (Spind4)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                            }
                            else if (Math.Floor(Netto) == 3 && Netto != 3)
                            {
                                if (!Spind1)
                                {
                                    writer.WriteLine("m101");
                                    Spind1 = true;
                                }

                                if (!Spind2)
                                {
                                    writer.WriteLine("m102");
                                    Spind2 = true;
                                }

                                if (!Spind3)
                                {
                                    writer.WriteLine("m103");
                                    Spind3 = true;
                                }

                                if (!Spind4)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = true;
                                }
                            }
                        }

                        if (Netto > 4 && Netto <= 8)
                        {
                            if (!Spind1)
                            {
                                writer.WriteLine("m101");
                                Spind1 = true;
                            }
                            if (!Spind2)
                            {
                                writer.WriteLine("m102");
                                Spind2 = true;
                            }
                            if (!Spind3)
                            {
                                writer.WriteLine("m103");
                                Spind3 = true;
                            }
                            if (!Spind4)
                            {
                                writer.WriteLine("m104");
                                Spind4 = true;
                            }
                        }

                        writer.WriteLine("G43 Z" + (radExt + secur) + " H1");
                        writer.WriteLine("G98 G81 X" + xErste + " Z" + (radInt - durch) + " R" + (radExt + secur) + " F9500");

                        writer.WriteLine("(----------------------------Form " + j + "-----------------------)");
                        writer.WriteLine("(offset " + offset + ")");

                        if (Netto <= 4)
                        {
                            if (Netto == Math.Floor(Netto))                 //wenn gerade zahl (Note: VB Int() equivalent is Math.Floor())
                            {
                                /*if (Netto == 1)                             //richtige spindeln auswählen
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind2 = false;
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Netto == 2)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Netto == 3)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }*/

                                writer.WriteLine("(S1:" + Spind1 + " S2:" + Spind2 + " S3:" + Spind3 + " S4:" + Spind4 + ")");

                                for (h = 1; h <= RPD; h++)                  //einmal bis 15, also ganzer durchgang (VB: For h = 1 To RPD)
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + (xPos + offset) + " A" + deg);
                                    for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                    {
                                        if (i == 1)
                                        {
                                            aPos = deg;
                                        }

                                        else if (i > 1)
                                        {
                                            aPos = aPos + deg;
                                            writer.WriteLine("A" + aPos + "");
                                        }
                                    }
                                    if (offsetbtn == true)
                                    {
                                        writer.WriteLine("(Versetzt)");
                                        //xPos = xErste + (h - 1) * (lab / 2);
                                        aPos = aNull;
                                        writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                        writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2));
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i == 2)
                                            {
                                                aPos = aPos + (deg / 2);
                                                writer.WriteLine("A" + aPos);
                                            }
                                            else if (i > 2)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos);
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                writer.WriteLine("(Not Whole)");            //wenn ungerade
                                /*if (Math.Floor(Netto) == 0)                 // VB: Int(Netto)
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind2 = false;
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 1)            // VB: Int(Netto)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 2)            // VB: Int(Netto)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }*/

                                writer.WriteLine("(S1:" + Spind1 + " S2:" + Spind2 + " S3:" + Spind3 + " S4:" + Spind4 + ")");
                                for (h = 1; h <= (int)Math.Floor(LastL); h++) //zuerst mit allen möglichen bis zum letzten punkt..... (VB: For h = 1 To Int(LastL))
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + (xPos + offset) + " A" + deg);
                                    Currentx = xPos;
                                    for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                    {
                                        if (i == 1)
                                        {
                                            aPos = deg;
                                        }
                                        if (i > 1)
                                        {
                                            aPos = aPos + deg;
                                            writer.WriteLine("A" + aPos + "");
                                        }
                                    }
                                    if (offsetbtn == true)
                                    {
                                        writer.WriteLine("(Versetzt)");
                                        //xPos = xErste + (h - 1) * (lab / 2);
                                        aPos = aNull;
                                        writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                        writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2));
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i == 2)
                                            {
                                                aPos = aPos + (deg / 2);
                                                writer.WriteLine("A" + aPos);
                                            }
                                            else if (i > 2)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos);
                                            }
                                        }
                                    }
                                }
                                if (Math.Floor(Netto) == 1)                 //die richtigen spindeln hochnehmen (VB: Int(Netto))
                                {
                                    writer.WriteLine("m102");
                                    Spind2 = false;
                                }
                                else if (Math.Floor(Netto) == 2)            // VB: Int(Netto)
                                {
                                    writer.WriteLine("m103");
                                    Spind3 = false;
                                }
                                else if (Math.Floor(Netto) == 3)            // VB: Int(Netto)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }

                                if (Netto >= 1)
                                {
                                    writer.WriteLine("(Rest nach spindelrückzug)");
                                    xPos = Currentx;
                                    // VB: For h = Int(LastL) To RPD - 1
                                    // C#: Loop runs from floor(LastL) up to and including RPD - 1
                                    for (h = (int)Math.Floor(LastL); h <= RPD - 1; h++) //den rest weiter bohren mit einer spindel weniger
                                    {
                                        xPos = xPos + lab;
                                        aPos = aNull;
                                        writer.WriteLine("(Reihe: " + (h + 1) + ")"); // Note: Original VB comment implies h starts from Int(LastL) and goes up to RPD-1. The row number is h+1.
                                        writer.WriteLine("X" + (xPos + offset) + " A" + deg);
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i > 1)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos + "");
                                            }
                                        }
                                        if (offsetbtn == true)
                                        {
                                            writer.WriteLine("(Versetzt)");
                                            //xPos = xErste + (h - 1) * (lab / 2);
                                            aPos = aNull;
                                            writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                            writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2)); ;
                                            for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                            {
                                                if (i == 1)
                                                {
                                                    aPos = deg;
                                                }
                                                if (i == 2)
                                                {
                                                    aPos = aPos + (deg / 2);
                                                    writer.WriteLine("A" + aPos);
                                                }
                                                else if (i > 2)
                                                {
                                                    aPos = aPos + deg;
                                                    writer.WriteLine("A" + aPos);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (Netto <= 8) // falls netto über 4 ist oder form über 236 lang
                        {
                            writer.WriteLine("(grösser 4)");
                            if (Netto == Math.Floor(Netto)) //gerade netto zahlen (VB: Int(Netto))
                            {
                                writer.WriteLine("(Gerade)");
                                for (h = 1; h <= RPD; h++) //einmal ganz durch fahren da es ja einer form über 4 ist (VB: For h = 1 To RPD)
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + (xPos + offset) + " A" + deg);
                                    for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                    {
                                        if (i == 1)
                                        {
                                            aPos = deg;
                                        }
                                        if (i > 1)
                                        {
                                            aPos = aPos + deg;
                                            writer.WriteLine("A" + aPos + "");
                                        }
                                    }
                                    if (offsetbtn == true)
                                    {
                                        writer.WriteLine("(Versetzt)");
                                        //xPos = xErste + (h - 1) * (lab / 2);
                                        aPos = aNull;
                                        writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                        writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2)); ;
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i == 2)
                                            {
                                                aPos = aPos + (deg / 2);
                                                writer.WriteLine("A" + aPos);
                                            }
                                            else if (i > 2)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos);
                                            }
                                        }
                                    }
                                }
                                if (Netto == 5)      //bohrer auswählen
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                }
                                else if (Netto == 6)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                }
                                else if (Netto == 7)
                                {
                                    writer.WriteLine("m104");
                                }
                                else if (Netto == 8)
                                {
                                    writer.WriteLine("(No Retraction!)");
                                }

                                for (h = 1; h <= RPD; h++) //mit ausgewählten bohrern ganz bohren (VB: For h = 1 To RPD)
                                {
                                    xPos = 240 + xErste + (h - 1) * lab; // Note the added 240 offset here
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + (xPos + offset) + " A" + deg);
                                    for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                    {
                                        if (i == 1)
                                        {
                                            aPos = deg;
                                        }
                                        if (i > 1)
                                        {
                                            aPos = aPos + deg;
                                            writer.WriteLine("A" + aPos + "");
                                        }
                                    }
                                    if (offsetbtn == true)
                                    {
                                        writer.WriteLine("(Versetzt)");
                                        //xPos = xErste + (h - 1) * (lab / 2);
                                        aPos = aNull;
                                        writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                        writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2)); ;
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i == 2)
                                            {
                                                aPos = aPos + (deg / 2);
                                                writer.WriteLine("A" + aPos);
                                            }
                                            else if (i > 2)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos);
                                            }
                                        }
                                    }
                                }
                            }
                            else //falls netto nicht gerade ist
                            {
                                writer.WriteLine("(ungerade)");
                                for (h = 1; h <= RPD; h++) //einmal ganz durch wie vorher (VB: For h = 1 To RPD)
                                {
                                    xPos = xErste + (h - 1) * lab;
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + (xPos + offset) + " A" + deg);

                                    for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                    {
                                        if (i == 1)
                                        {
                                            aPos = deg;
                                        }
                                        if (i > 1)
                                        {
                                            aPos = aPos + deg;
                                            writer.WriteLine("A" + aPos + "");
                                        }
                                    }
                                    if (offsetbtn == true)
                                    {
                                        writer.WriteLine("(Versetzt)");
                                        //xPos = xErste + (h - 1) * (lab / 2);
                                        aPos = aNull;
                                        writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                        writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2)); ;
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i == 2)
                                            {
                                                aPos = aPos + (deg / 2);
                                                writer.WriteLine("A" + aPos);
                                            }
                                            else if (i > 2)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos);
                                            }
                                        }
                                    }
                                }

                                writer.WriteLine("(Not Whole)"); //bohrer...
                                if (Math.Floor(Netto) == 4) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m102");
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind2 = false;
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 5) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m103");
                                    writer.WriteLine("m104");
                                    Spind3 = false;
                                    Spind4 = false;
                                }
                                else if (Math.Floor(Netto) == 6) // VB: Int(Netto)
                                {
                                    writer.WriteLine("m104");
                                    Spind4 = false;
                                }
                                writer.WriteLine("(S1:" + Spind1 + " S2:" + Spind2 + " S3:" + Spind3 + " S4:" + Spind4 + ")");
                                // VB: For h = 1 To Int(LastL) + 1
                                // C#: Loop runs from 1 up to and including floor(LastL) + 1
                                for (h = 1; h <= (int)Math.Floor(LastL); h++) //bis zur letzten bohren mit allen möglichen
                                {
                                    writer.WriteLine("(" + h + ")");
                                    writer.WriteLine("(" + (int)Math.Floor(LastL) + ")");
                                    xPos = 240 + xErste + (h - 1) * lab; // Note the added 240 offset here
                                    aPos = aNull;
                                    writer.WriteLine("(Reihe: " + h + ")");
                                    writer.WriteLine("X" + (xPos + offset) + " A" + deg);
                                    Currentx = xPos;
                                    for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                    {
                                        if (i == 1)
                                        {
                                            aPos = deg;
                                        }
                                        if (i > 1)
                                        {
                                            aPos = aPos + deg;
                                            writer.WriteLine("A" + aPos + "");
                                        }
                                    }
                                    if (offsetbtn == true)
                                    {
                                        writer.WriteLine("(Versetzt)");
                                        //xPos = xErste + (h - 1) * (lab / 2);
                                        aPos = aNull;
                                        writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                        writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2)); ;
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i == 2)
                                            {
                                                aPos = aPos + (deg / 2);
                                                writer.WriteLine("A" + aPos);
                                            }
                                            else if (i > 2)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos);
                                            }
                                        }
                                    }
                                }
                                if (Math.Floor(Netto) == 4) //Bohrer (VB: Int(Netto))
                                {
                                    if (Spind1 == true)
                                    {
                                        writer.WriteLine("m101");
                                        Spind1 = false;
                                    }
                                }
                                else if (Math.Floor(Netto) == 5) // VB: Int(Netto)
                                {
                                    if (Spind2 == true)
                                    {
                                        writer.WriteLine("m102");
                                        Spind2 = false;
                                    }
                                }
                                else if (Math.Floor(Netto) == 6) // VB: Int(Netto)
                                {
                                    if (Spind3 == true)
                                    {
                                        writer.WriteLine("m103");
                                        Spind3 = false;
                                    }
                                }
                                else if (Math.Floor(Netto) == 7) // VB: Int(Netto)
                                {
                                    if (Spind4 == true)
                                    {
                                        writer.WriteLine("m104");
                                        Spind4 = false;
                                    }
                                }
                                // spindelrückzug
                                if (Netto >= 5)
                                {
                                    writer.WriteLine("(Rest nach spindelrückzug)");
                                    xPos = Currentx;
                                    // VB: For h = Int(LastL) To RPD - 2
                                    // C#: Loop runs from floor(LastL) up to and including RPD - 2
                                    for (h = (int)Math.Floor(LastL); h <= RPD - 1; h++) //den rest auffüllen
                                    {
                                        xPos = xPos + lab;
                                        aPos = aNull;
                                        writer.WriteLine("(Reihe: " + (h + 1) + ")"); // Note: Row number is h+1
                                        writer.WriteLine("X" + (xPos + offset) + " A" + deg);
                                        for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                        {
                                            if (i == 1)
                                            {
                                                aPos = deg;
                                            }
                                            if (i > 1)
                                            {
                                                aPos = aPos + deg;
                                                writer.WriteLine("A" + aPos + "");
                                            }
                                        }
                                        if (offsetbtn == true)
                                        {
                                            writer.WriteLine("(Versetzt)");
                                            //xPos = xErste + (h - 1) * (lab / 2);
                                            aPos = aNull;
                                            writer.WriteLine("(Versetzte Reihe: " + h + ")");
                                            writer.WriteLine("X" + (xPos + (lab / 2) + offset) + " A" + (deg / 2)); ;
                                            for (i = 1; i <= nbrHolesRad; i++)      // VB: For i = 1 To nbrHolesRad
                                            {
                                                if (i == 1)
                                                {
                                                    aPos = deg;
                                                }
                                                if (i == 2)
                                                {
                                                    aPos = aPos + (deg / 2);
                                                    writer.WriteLine("A" + aPos);
                                                }
                                                else if (i > 2)
                                                {
                                                    aPos = aPos + deg;
                                                    writer.WriteLine("A" + aPos);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        writer.WriteLine("G80");
                    }

                    writer.WriteLine("m110"); //alle spindeln hoch
                    writer.WriteLine("G28 G91 Z0.");    //hiemfahrt z achse
                    writer.WriteLine("G90");            //absolute positionierung
                    writer.WriteLine("A0.");            //absolute winkelpositionierung
                    writer.WriteLine("G28 G91 X0.");    //hiemfahrt x achse
                    //writer.WriteLine("G90");
                    writer.WriteLine("M5");             //spindel stopp
                    writer.WriteLine("M30");            //programmende

                    writer.Close();
                }
                #endregion
            });
            #region -- speichern --

            var s = await _settingsService.LoadAsync();
            s.PathInput = directory;

            if (IsChecked == true)
            {
                s.Height = double.Parse(_height.Text ?? "0");
                s.Lab = double.Parse(_lab.Text ?? "0");
                s.Mat = double.Parse(_mat.Text ?? "0");
                s.XErste = double.Parse(_xErste.Text ?? "0");
                s.Rabo = double.Parse(_rabo.Text ?? "0");
                s.Deg = double.Parse(_deg.Text ?? "0");
                s.Durch = double.Parse(_durch.Text ?? "0");
                s.SpinRpm = _spinRPM.Value;
                s.XVersch = double.Parse(_xVersch.Text ?? "0");
                s.DiaInt = double.Parse(_diaInt.Text ?? "0");
                s.Secur = double.Parse(_secur.Text ?? "0");
                s.XAbst = double.Parse(_xAbst.Text ?? "0");
                s.Custom = _custom.IsChecked == true;
                s.Anz = double.Parse(_anz.Text ?? "0");

                s.Kunde = _firmaName.Text ?? "";

                s.KeepValues = true;
                s.Farbe = _farbe.IsChecked == true;
                s.Versetzt = _versetzt.IsChecked == true;

                string selectedMode = OptionA.IsChecked == true ? "0.8" :
                            OptionB.IsChecked == true ? "1.0" :
                            OptionC.IsChecked == true ? "1.3" :
                            OptionD.IsChecked == true ? "1.5" :
                            OptionE.IsChecked == true ? "2.0" :
                            OptionF.IsChecked == true ? "3.0" :
                            "None";

                s.DrillDia = selectedMode;
            }
            else
            {
                s.KeepValues = false;
                ClearInputs(null, new RoutedEventArgs());
            }

            await _settingsService.SaveAsync(s);
            #endregion

            await ShowInfoAsync(this, "Erfolgreiche Eingabe", "Code Erstellt unter: " + filePath + "\nZ = " + (radInt - durch) + " | R = " + (radExt + secur));
        }
        catch (Exception ex)
        {
            await ShowErrorAsync(this, "Error", $"Error: {ex.Message}");
        }
        #endregion
    }
}
