Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Crear el formulario
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mapeo de PST en Outlook"
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Label para mostrar ruta
$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Location = New-Object System.Drawing.Point(10,20)
$labelPath.Size = New-Object System.Drawing.Size(360,20)
$labelPath.Text = "Carpeta seleccionada: (ninguna)"
$form.Controls.Add($labelPath)

# Botón para abrir el explorador de carpetas
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(10,50)
$btnBrowse.Size = New-Object System.Drawing.Size(150,30)
$btnBrowse.Text = "Seleccionar carpeta"
$btnBrowse.BackColor = [System.Drawing.Color]::LightBlue
$form.Controls.Add($btnBrowse)

# Botón para mapear PST
$btnMap = New-Object System.Windows.Forms.Button
$btnMap.Location = New-Object System.Drawing.Point(200,50)
$btnMap.Size = New-Object System.Drawing.Size(150,30)
$btnMap.Text = "Mapear PST"
$btnMap.BackColor = [System.Drawing.Color]::LightBlue
$btnMap.Enabled = $false
$form.Controls.Add($btnMap)

# Label de estado
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10,100)
$labelStatus.Size = New-Object System.Drawing.Size(360,50)
$labelStatus.Text = ""
$form.Controls.Add($labelStatus)

# Evento botón Seleccionar carpeta
$btnBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Selecciona la carpeta que contiene los archivos PST"
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderBrowser.SelectedPath
        $labelPath.Text = "Carpeta seleccionada: $selectedFolder"
        $btnMap.Enabled = $true
        $labelStatus.Text = ""
    }
})

# Función para mapear PST
function MapPST {
    param ($folderPath)

    # Limpiar label status
    $labelStatus.Text = "Iniciando proceso..."

    # Verificar que existan PST
    $pstFiles = Get-ChildItem -Path $folderPath -Filter *.pst -ErrorAction SilentlyContinue
    if ($pstFiles.Count -eq 0) {
        $labelStatus.Text = "No se encontraron archivos .pst en la carpeta."
        return
    }

    # Intentar iniciar Outlook
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
    } catch {
        $labelStatus.Text = "No se pudo iniciar Outlook. Verifica que esté instalado."
        return
    }

    # Obtener PSTs ya montados para evitar duplicados
    $mountedStores = @()
    for ($i = 1; $i -le $namespace.Folders.Count; $i++) {
        $store = $namespace.Folders.Item($i)
        try {
            $filePath = $store.Store.FilePath
            if ($filePath) {
                $mountedStores += $filePath.ToLower()
            }
        } catch {}
    }

    $mappedCount = 0
    $skippedCount = 0

    foreach ($pst in $pstFiles) {
        $pstPathLower = $pst.FullName.ToLower()
        if ($mountedStores -contains $pstPathLower) {
            $skippedCount++
        } else {
            try {
                $namespace.AddStore($pst.FullName)
                $mappedCount++
            } catch {
                $labelStatus.Text = "Error cargando $($pst.Name): $($_.Exception.Message)"
                return
            }
        }
    }

    $labelStatus.Text = "Proceso terminado.`nMapeados: $mappedCount .pst(s). Omitidos (ya cargados): $skippedCount."

}

# Evento botón Mapear PST corregido para usar la ruta del label
$btnMap.Add_Click({
    # Obtener la carpeta directamente del label para evitar desincronización
    $pathLabelText = $labelPath.Text
    # Extraer la ruta (quitando el prefijo "Carpeta seleccionada: ")
    $folder = $pathLabelText -replace "^Carpeta seleccionada: ", ""

    if (-not (Test-Path -Path $folder)) {
        $labelStatus.Text = "Ruta no válida. Selecciona una carpeta válida."
        return
    }

    MapPST -folderPath $folder

    # Reiniciar interfaz para nuevo uso
    $labelPath.Text = "Carpeta seleccionada: (ninguna)"
    $btnMap.Enabled = $false
})

# Mostrar el formulario
[void]$form.ShowDialog()
