' ============================================================
'  Apagado.vbs  —  Auto-instalador + Lanzador
'  Contiene el script PowerShell embebido.
'  La primera vez que se ejecuta se copia a shell:startup.
'  Las veces siguientes simplemente lanza el temporizador.
' ============================================================

Option Explicit

Dim WshShell, FSO, StartupFolder, ThisFile, StartupFile

Set WshShell = CreateObject("WScript.Shell")
Set FSO      = CreateObject("Scripting.FileSystemObject")

StartupFolder = WshShell.SpecialFolders("Startup")
ThisFile      = WScript.ScriptFullName
StartupFile   = StartupFolder & "\Apagado.vbs"

If LCase(ThisFile) <> LCase(StartupFile) Then
    FSO.CopyFile ThisFile, StartupFile, True
    MsgBox "Instalado correctamente en Inicio de Windows." & vbCrLf & _
           "El temporizador arrancara automaticamente con cada encendido.", _
           vbInformation, "Apagado - Instalacion completa"
End If

Dim PS1Code
PS1Code = _
"Add-Type -AssemblyName PresentationFramework, System.Drawing, WindowsBase" & vbCrLf & _
"" & vbCrLf & _
"$Global:HoraFinal = ''" & vbCrLf & _
"$Global:SegundosRestantes = 0" & vbCrLf & _
"$Global:AvisoMostrado = $false" & vbCrLf & _
"" & vbCrLf & _
"while ($true) {" & vbCrLf & _
"    $xmlConfig = @'" & Chr(10) & _
"<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'" & Chr(10) & _
"        Title='Config' Height='220' Width='300' WindowStartupLocation='CenterScreen'" & Chr(10) & _
"        WindowStyle='None' AllowsTransparency='True' Background='Transparent' Topmost='True'>" & Chr(10) & _
"    <Border BorderBrush='#444' BorderThickness='2' CornerRadius='15' Background='#F2111111'>" & Chr(10) & _
"        <StackPanel VerticalAlignment='Center' HorizontalAlignment='Center'>" & Chr(10) & _
"            <TextBlock Text='HORA DE APAGADO' Foreground='Yellow' FontSize='16' FontWeight='Bold' HorizontalAlignment='Center' Margin='10'/>" & Chr(10) & _
"            <TextBox Name='InputHora' Text='17:00' FontSize='20' Width='100' TextAlignment='Center' Background='#333' Foreground='White' BorderThickness='0' Margin='5'/>" & Chr(10) & _
"            <Button Name='BtnIniciar' Content='Iniciar Jornada' Width='130' Height='35' Margin='15' Background='#007ACC' Foreground='White' FontWeight='Bold'/>" & Chr(10) & _
"        </StackPanel>" & Chr(10) & _
"    </Border>" & Chr(10) & _
"</Window>" & Chr(10) & _
"'@" & vbCrLf & _
"    $readerConfig = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xmlConfig)" & vbCrLf & _
"    $VentanaConfig = [System.Windows.Markup.XamlReader]::Load($readerConfig)" & vbCrLf & _
"    $InputHora = $VentanaConfig.FindName('InputHora')" & vbCrLf & _
"    $VentanaConfig.FindName('BtnIniciar').Add_Click({" & vbCrLf & _
"        $Global:HoraFinal = $InputHora.Text" & vbCrLf & _
"        $VentanaConfig.Close()" & vbCrLf & _
"    })" & vbCrLf & _
"    $VentanaConfig.ShowDialog() | Out-Null" & vbCrLf & _
"    if ($Global:HoraFinal -eq '') { exit }" & vbCrLf & _
"    try {" & vbCrLf & _
"        $Ahora = Get-Date" & vbCrLf & _
"        $Destino = Get-Date $Global:HoraFinal" & vbCrLf & _
"        if ($Destino -lt $Ahora) { $Destino = $Destino.AddDays(1) }" & vbCrLf & _
"        $Global:SegundosRestantes = [Math]::Floor(($Destino - $Ahora).TotalSeconds)" & vbCrLf & _
"        break" & vbCrLf & _
"    } catch {" & vbCrLf & _
"        [System.Windows.MessageBox]::Show('Formato incorrecto. Usa HH:mm (ej: 17:30)', 'Error')" & vbCrLf & _
"        $Global:HoraFinal = ''" & vbCrLf & _
"    }" & vbCrLf & _
"}" & vbCrLf & _
"" & vbCrLf & _
"$xmlReloj = @'" & Chr(10) & _
"<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'" & Chr(10) & _
"        Title='Reloj' Height='70' Width='230' AllowsTransparency='True'" & Chr(10) & _
"        WindowStyle='None' Background='Transparent' Topmost='True' ShowInTaskbar='False'>" & Chr(10) & _
"    <StackPanel VerticalAlignment='Center'>" & Chr(10) & _
"        <TextBlock Name='RelojText' Text='00:00:00' FontSize='26' Foreground='White' Opacity='0.45' FontWeight='Bold' TextAlignment='Left' Margin='15,0,0,2'>" & Chr(10) & _
"            <TextBlock.Effect>" & Chr(10) & _
"                <DropShadowEffect BlurRadius='10' ShadowDepth='4' Color='Black' Opacity='0.9'/>" & Chr(10) & _
"            </TextBlock.Effect>" & Chr(10) & _
"        </TextBlock>" & Chr(10) & _
"        <TextBlock Name='SubText' FontSize='11' Foreground='White' Opacity='0.45' TextAlignment='Left' Margin='15,0,0,0'>" & Chr(10) & _
"            <TextBlock.Effect>" & Chr(10) & _
"                <DropShadowEffect BlurRadius='5' ShadowDepth='2' Color='Black' Opacity='0.8'/>" & Chr(10) & _
"            </TextBlock.Effect>" & Chr(10) & _
"        </TextBlock>" & Chr(10) & _
"    </StackPanel>" & Chr(10) & _
"</Window>" & Chr(10) & _
"'@" & vbCrLf & _
"$readerReloj = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xmlReloj)" & vbCrLf & _
"$WindowReloj = [System.Windows.Markup.XamlReader]::Load($readerReloj)" & vbCrLf & _
"$RelojLabel = $WindowReloj.FindName('RelojText')" & vbCrLf & _
"$SubLabel   = $WindowReloj.FindName('SubText')" & vbCrLf & _
"$SubLabel.Text = 'Apagado: ' + $Global:HoraFinal" & vbCrLf & _
"$WindowReloj.Left = 10" & vbCrLf & _
"$WindowReloj.Top  = [System.Windows.SystemParameters]::PrimaryScreenHeight - 180" & vbCrLf & _
"" & vbCrLf & _
"$xmlAviso = @'" & Chr(10) & _
"<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'" & Chr(10) & _
"        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'" & Chr(10) & _
"        Title='Aviso' Height='230' Width='400' WindowStartupLocation='CenterScreen'" & Chr(10) & _
"        WindowStyle='None' AllowsTransparency='True' Background='Transparent' Topmost='True' Visibility='Hidden'>" & Chr(10) & _
"    <Border BorderBrush='#555' BorderThickness='1' CornerRadius='15' Background='#F2111111'>" & Chr(10) & _
"        <StackPanel VerticalAlignment='Center' HorizontalAlignment='Center' Margin='20'>" & Chr(10) & _
"            <TextBlock Text='FALTAN 5 MINUTOS' Foreground='#F5C518' FontSize='20' FontWeight='Bold' HorizontalAlignment='Center' Margin='0,0,0,6'/>" & Chr(10) & _
"            <TextBlock Text='Que queres hacer?' Foreground='#AAAAAA' FontSize='13' HorizontalAlignment='Center' Margin='0,0,0,20'/>" & Chr(10) & _
"            <UniformGrid Rows='2' Columns='2'>" & Chr(10) & _
"                <Button Name='BtnExtra5'   Content='+5 min'          Height='42' Margin='5' Background='#1a5276' Foreground='White' FontWeight='Bold' BorderThickness='0'/>" & Chr(10) & _
"                <Button Name='BtnSeguir'   Content='Continuar'       Height='42' Margin='5' Background='#1e8449' Foreground='White' FontWeight='Bold' BorderThickness='0'/>" & Chr(10) & _
"                <Button Name='BtnCancelar' Content='Cancelar apagado' Height='42' Margin='5' Background='#444444' Foreground='White' FontWeight='Bold' BorderThickness='0'/>" & Chr(10) & _
"                <Button Name='BtnAhora'    Content='Apagar ahora'    Height='42' Margin='5' Background='#922b21' Foreground='White' FontWeight='Bold' BorderThickness='0'/>" & Chr(10) & _
"            </UniformGrid>" & Chr(10) & _
"        </StackPanel>" & Chr(10) & _
"    </Border>" & Chr(10) & _
"</Window>" & Chr(10) & _
"'@" & vbCrLf & _
"$readerAviso = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xmlAviso)" & vbCrLf & _
"$VentanaAviso = [System.Windows.Markup.XamlReader]::Load($readerAviso)" & vbCrLf & _
"" & vbCrLf & _
"$VentanaAviso.FindName('BtnExtra5').Add_Click({" & vbCrLf & _
"    $Global:SegundosRestantes += 300" & vbCrLf & _
"    $Global:AvisoMostrado = $false" & vbCrLf & _
"    $VentanaAviso.Visibility = 'Hidden'" & vbCrLf & _
"})" & vbCrLf & _
"$VentanaAviso.FindName('BtnSeguir').Add_Click({" & vbCrLf & _
"    $VentanaAviso.Visibility = 'Hidden'" & vbCrLf & _
"})" & vbCrLf & _
"$VentanaAviso.FindName('BtnCancelar').Add_Click({" & vbCrLf & _
"    $timer.Stop()" & vbCrLf & _
"    $WindowReloj.Close()" & vbCrLf & _
"    $VentanaAviso.Close()" & vbCrLf & _
"    exit" & vbCrLf & _
"})" & vbCrLf & _
"$VentanaAviso.FindName('BtnAhora').Add_Click({" & vbCrLf & _
"    shutdown /s /f /t 0" & vbCrLf & _
"})" & vbCrLf & _
"" & vbCrLf & _
"$timer = New-Object System.Windows.Threading.DispatcherTimer" & vbCrLf & _
"$timer.Interval = [TimeSpan]::FromSeconds(1)" & vbCrLf & _
"$timer.Add_Tick({" & vbCrLf & _
"    $Global:SegundosRestantes--" & vbCrLf & _
"    $seg = [Math]::Max(0, $Global:SegundosRestantes)" & vbCrLf & _
"    $t = [Timespan]::FromSeconds($seg)" & vbCrLf & _
"    $RelojLabel.Text = '{0:D2}:{1:D2}:{2:D2}' -f $t.Hours, $t.Minutes, $t.Seconds" & vbCrLf & _
"    if ($Global:SegundosRestantes -le 300) {" & vbCrLf & _
"        $RelojLabel.Foreground = 'Red'" & vbCrLf & _
"    } elseif ($Global:SegundosRestantes -le 1800) {" & vbCrLf & _
"        $RelojLabel.Foreground = 'Orange'" & vbCrLf & _
"    } else {" & vbCrLf & _
"        $RelojLabel.Foreground = 'White'" & vbCrLf & _
"    }" & vbCrLf & _
"    if ($Global:SegundosRestantes -le 300 -and -not $Global:AvisoMostrado) {" & vbCrLf & _
"        $Global:AvisoMostrado = $true" & vbCrLf & _
"        $VentanaAviso.Visibility = 'Visible'" & vbCrLf & _
"    }" & vbCrLf & _
"    if ($Global:SegundosRestantes -le 0) {" & vbCrLf & _
"        $timer.Stop()" & vbCrLf & _
"        shutdown /s /f /t 0" & vbCrLf & _
"    }" & vbCrLf & _
"})" & vbCrLf & _
"" & vbCrLf & _
"$timer.Start()" & vbCrLf & _
"$WindowReloj.ShowDialog() | Out-Null"

Dim TempPS1
TempPS1 = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\ApagadoTimer.ps1"

Dim FileOut
Set FileOut = FSO.CreateTextFile(TempPS1, True, True)
FileOut.Write PS1Code
FileOut.Close

WshShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & TempPS1 & """", 0, False
