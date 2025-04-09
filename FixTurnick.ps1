Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
 
# Ocultar ventana de PowerShell
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class WinAPI {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@
$pswindow = (Get-Process -Id $pid).MainWindowHandle
[WinAPI]::ShowWindow($pswindow, 0)
 
# Crear formulario principal
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "FixTurnick"
$Form.Size = New-Object System.Drawing.Size(900, 600)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::black
$Form.Font = New-Object System.Drawing.Font("Serif", 10)
 
# Cuadro de texto para logs
$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.Size = New-Object System.Drawing.Size(300, 400) 
$LogBox.Location = New-Object System.Drawing.Point(450, 30)#30 fue subir #450 fue la ubicacion de lado
$LogBox.ReadOnly = $true
$LogBox.BackColor = [System.Drawing.Color]::White
 
$LogBox.ScrollBars = "Vertical"
$Form.Controls.Add($LogBox)
 
Function Add-Log {
    param ([string]$Message)
    $LogBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $Message`r`n")

}
    
 
# Panel de botones
$Panel = New-Object System.Windows.Forms.Panel
$Panel.Size = New-Object System.Drawing.Size(440, 700)
$Panel.Location = New-Object System.Drawing.Point(20, 20)
$Form.Controls.Add($Panel)
 
Function Create-Button {
    param ($Text, $Location, $Action)
    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = $Text
    $Button.Size = New-Object System.Drawing.Size(200, 40)
    $Button.Location = $Location
    $Button.BackColor = [System.Drawing.Color]::White
    $Button.Add_Click($Action)
    $Panel.Controls.Add($Button)
}

 
# Definir botones
Create-Button "Clean Temp Files" (New-Object System.Drawing.Point(10, 10)) {
    Add-Log "Cleaning temporary files..."
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c del /s /q %temp%\*"
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c del /s /q C:\Windows\Temp\*"
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c del /s /q C:\Users\xhclasanchez\AppData\Roaming\Microsoft\*"
    Add-Log "Temporary files cleaned."
}

Create-Button "Repair System" (New-Object System.Drawing.Point(220, 10)) {
    Add-Log "Running system repair..."
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c DISM /Online /Cleanup-Image /checkhealth"
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c defrag"
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c sfc /scannow"
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c DISM /online /cleanup-image /restorehealth"
    Add-Log "System repair completed."
}

Create-Button "Panel de Control" (New-Object System.Drawing.Point(10, 60)) {
   Start-Process -FilePath "C:\Windows\System32\control.exe"
  
}

Create-Button "Virus scanner" (New-Object System.Drawing.Point(220, 60)) {
    Add-Log "Scan for viruses"
   Start-Process -FilePath "C:\Windows\System32\MRT.exe"
    Add-Log "Scanning completed."
}

Create-Button "Windows Defender" (New-Object System.Drawing.Point(10, 110)) {
    Add-Log "MRT"
  Start-Process "windowsdefender:"
    Add-Log "AntiVirus."
}

Create-Button "Check Windows Updates" (New-Object System.Drawing.Point(10, 160)) {
    Add-Log "Checking for Windows updates..."

    Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -Command `"Import-Module PSWindowsUpdate; Get-WindowsUpdate -AcceptAll -Install -AutoReboot`"" -NoNewWindow -Wait

    Add-Log "Windows update check initiated."
}

Create-Button "Check Software Updates" (New-Object System.Drawing.Point(10, 210)) {
     Add-Log "Checking for software updates..."
 Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd -ArgumentList '/c winget upgrade --all' -Verb RunAs" -NoNewWindow

 Add-Log "Windows update check"
}

Create-Button "Force Group Policy Update" (New-Object System.Drawing.Point(220, 110)) {
    Add-Log "Updating group policies..."
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force"
    Add-Log "Group policies updated."
}

Create-Button "Dell Drivers Page" (New-Object System.Drawing.Point(10, 260)) {
    Add-Log "Opening Dell drivers page..."
    Start-Process "https://www.dell.com/support/home/pt-br?app=drivers"
}

Create-Button "Lenovo Drivers Page" (New-Object System.Drawing.Point(10, 310)) {
    Add-Log "Opening Lenovo drivers page..."
    Start-Process "https://www.dell.com/support/home/pt-br?app=drivers"
}

Create-Button "Restart System" (New-Object System.Drawing.Point(10, 360)) {
    Add-Log "Restarting system..."
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c shutdown /r /t 10"
}
Create-Button "Reset Network" (New-Object System.Drawing.Point(10, 410)) {
    Add-Log "Resetting network..."
    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c ipconfig /flushdns & netsh winsock reset"
    Add-Log "Network reset completed."
}

Create-Button "Activate Nuclei" (New-Object System.Drawing.Point(220, 160)) {
    Add-Log "Activating all CPU cores..."
    Start-Process -FilePath "bcdedit.exe" -ArgumentList "/deletevalue {current} numproc" -Verb RunAs
    Add-Log "All CPU cores activated. Restart required."
}

Create-Button "Resource Consumption" (New-Object System.Drawing.Point(220, 210)) {
    Add-Log "Checking processes that are consuming more resources..."

    # Crear la nueva ventana
    $Form3 = New-Object System.Windows.Forms.Form
    $Form3 = New-Object System.Windows.Forms.Form
    $Form3.Text = "Resource Consumption"
    $Form3.Size = New-Object System.Drawing.Size(600, 500)
    $Form.BackColor = ::WhiteSmoke
   
  # Agregar columnas
    $Form3.Columns.Add("Status", 100) | Out-Null
    $Form3.Columns.Add("Name", 150) | Out-Null
    # Crear un TextBox para mostrar los procesos
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Multiline = $true
    $TextBox.ScrollBars = "Vertical"
    $TextBox.Size = New-Object System.Drawing.Size(600, 500)
    $TextBox.Location = New-Object System.Drawing.Point(10, 10)
    

    # Obtener procesos y asignarlos al TextBox
    $processes = Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 
    $TextBox.Text = ($processes | Out-String)
   

    # Agregar el TextBox al formulario y mostrarlo
    $Form3.Controls.Add($TextBox)
    $Form3.ShowDialog()
}
# Mostrar la ventana


Create-Button "List of Running Services" (New-Object System.Drawing.Point(220, 260)) {
    Add-Log "Checking background services.."
    # Crear una nueva ventana personalizada
    $Form2 = New-Object System.Windows.Forms.Form
    $Form2.Text = "List of Running Services"
    $Form2.Size = New-Object System.Drawing.Size(600, 500)wx
    $Form2.BackColor = [System.Drawing.Color]::WhiteSmoke
    # Crear ListView personalizado
    $ListView = New-Object System.Windows.Forms.ListView
    $ListView.View = "Details"
    $ListView.FullRowSelect = $true
    $ListView.GridLines = $true
    $ListView.Size = New-Object System.Drawing.Size(1000, 900)
    $ListView.Location = New-Object System.Drawing.Point(10, 10)
    # Agregar columnas
    $ListView.Columns.Add("Status", 100) | Out-Null
    $ListView.Columns.Add("Name", 150) | Out-Null
    $ListView.Columns.Add("DisplayName", 300) | Out-Null
    # Agregar elementos a la lista
    Get-Service | Where-Object { $_.Status -eq "Running" } | ForEach-Object {
        $Item = New-Object System.Windows.Forms.ListViewItem
        $Item.Text = $_.Status
        $Item.SubItems.Add($_.Name) | Out-Null
        $Item.SubItems.Add($_.DisplayName) | Out-Null
        # Cambiar color de elementos (ejemplo: resaltar servicios específicos)
        if ($_.Name -like "*ProtonVPN*") {
            $Item.BackColor = [System.Drawing.Color]::LightYellow
        }
        $ListView.Items.Add($Item) | Out-Null
    }
    # Añadir la lista al formulario
    $Form2.Controls.Add($ListView)
    $Form2.ShowDialog()

}

# Botón para exportar el reporte
 Create-Button "Export Report to PDF" (New-Object System.Drawing.Point(220, 310)) {
    Add-Log "Generating report..."

    # Definir rutas de archivos
    $docxPath = "$env:TEMP\SystemReport.docx"
    $pdfPath = "$env:TEMP\SystemReport.pdf"

    # Capturar el contenido del LogBox
    $logContent = $LogBox.Text

    # Verificar si hay contenido en el LogBox
    if ([string]::IsNullOrWhiteSpace($logContent)) {
        [System.Windows.Forms.MessageBox]::Show("No data to export!", "Warning", "OK", "Warning")
        return
    }

    # Verificar si Microsoft Word está instalado
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false

        # Crear un nuevo documento de Word
        $doc = $word.Documents.Add()

        # Seleccionar el contenido del documento y escribir el contenido del LogBox
        $selection = $word.Selection
        $selection.TypeText($logContent)

        # Guardar como DOCX
        $doc.SaveAs([ref]$docxPath, [ref]16)  # 16 = formato DOCX

        # Guardar como PDF
        $pdfFormat = 17  # 17 = formato PDF
        $doc.SaveAs([ref]$pdfPath, [ref]$pdfFormat)

        # Cerrar Word
        $doc.Close()
        $word.Quit()

        Add-Log "Report successfully saved as PDF: $pdfPath"
        [System.Windows.Forms.MessageBox]::Show("Report saved to $pdfPath", "Success", "OK", "Information")

        # Abrir el PDF automáticamente
        Start-Process $pdfPath
    }
    catch {
        Add-Log "Error: Microsoft Word is not installed or failed to convert file."
        [System.Windows.Forms.MessageBox]::Show("Error: Microsoft Word is not installed or failed to generate the PDF.", "Error", "OK", "Error")
    }
}



$Form.ShowDialog()
