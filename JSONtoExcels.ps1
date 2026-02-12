#Author: Santiago Alvarez
#Githb: SanntAlv

$baseDir = $PSScriath "resultado.txt"
$directorioReportesJson = Join-Path -Path $baseDir -ChildPath "devops-powershell\reportes" 
$directorioExcelFinal   = Join-Path -Path $baseDir -ChildPath "devops-powershell\reportes\proactiva-excel"
$rutaArchivoResultado = Join-Path -Path $baseDir -ChildP
try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Error "El módulo 'ImportExcel' no está instalado en PowerShell portable. Por favor descargue el modulo y peguelo en el directorio \devops-powershell-portable 1.7.4()\Modules."
    Read-Host "Presiona Enter para salir."
    exit 1
}

Write-Host "Iniciando conversión de JSON a Excel..." -ForegroundColor Green

# Verificamos si existe la lista de tareas
if (-not (Test-Path $rutaArchivoResultado)) {
    Write-Warning "No se encontró el archivo de resultado 'resultado.txt'. No hay archivos para procesar."
    exit 0 # Salimos sin error porque no hay nada que hacer
}

$listaDeArchivosJson = Get-Content -Path $rutaArchivoResultado

foreach ($nombreJson in $listaDeArchivosJson) {
    try {
        $rutaJsonEntrada = Join-Path -Path $directorioReportesJson -ChildPath $nombreJson
        if (-not (Test-Path $rutaJsonEntrada)) {
            Write-Warning "El archivo '$nombreJson' listado en resultado.txt no fue encontrado en la carpeta de reportes. Omitiendo."
            continue # Salta al siguiente archivo de la lista
        }

        Write-Host "`nProcesando archivo: '$nombreJson'..." -ForegroundColor Yellow
        
        # --- (Toda tu lógica de conversión ahora va DENTRO del bucle) ---
        $jsonData = Get-Content -Path $rutaJsonEntrada -Raw | ConvertFrom-Json
        $reportObject = $jsonData.Report
        
        $nombreExcel = $nombreJson.Replace(".json", ".xlsx")
        $rutaExcelFinalCompleta = Join-Path -Path $directorioExcelFinal -ChildPath $nombreExcel
        
        if (Test-Path $rutaExcelFinalCompleta) {
            Remove-Item $rutaExcelFinalCompleta
        }

        Write-Host "Exportando datos a: $rutaExcelFinalCompleta"
        foreach ($sheet in $reportObject.PSObject.Properties) {
            if ($sheet.Value -and $sheet.Value.Count -gt 0) {
                Write-Host "  -> Creando hoja: '$($sheet.Name)'..."
                $exportParams = @{
                    Path          = $rutaExcelFinalCompleta
                    WorksheetName = $sheet.Name
                    AutoSize      = $true
                    FreezeTopRow  = $true
                    AutoFilter    = $true
                }

                if ($sheet.Name -eq 'ESXi') { $exportParams['NoNumberConversion'] = @('NtpServer', 'DnsServer', 'Hostname') }
                if ($sheet.Name -eq 'vCenter') { $exportParams['NoNumberConversion'] = @('Version') }
                if ($sheet.Name -eq 'VM') { $exportParams['NoNumberConversion'] = @('Host') }
                if ($sheet.Name -eq 'Datastores') { $exportParams['NoNumberConversion'] = @('Hostname') }
                if ($sheet.Name -eq 'Standard Switch') { $exportParams['NoNumberConversion'] = @('ESXi') }
                if ($sheet.Name -eq 'VMkernel Adapters') { $exportParams['NoNumberConversion'] = @('Host', 'IP') }
                if ($sheet.Name -eq 'vNetwork') { $exportParams['NoNumberConversion'] = @('Host') }
                if ($sheet.Name -eq 'Falso Positivo') { $exportParams['NoNumberConversion'] = @('Host') }
                if ($sheet.Name -eq 'Compatibilidad de Componentes') { $exportParams['NoNumberConversion'] = @('Hostname') }

                $sheet.Value | Export-Excel @exportParams
            }
        }

        Write-Host "Aplicando estilos al archivo Excel..."
        $excelPackage = Open-ExcelPackage -Path $rutaExcelFinalCompleta
        foreach ($ws in $excelPackage.Workbook.Worksheets) {
            $headerRange = $ws.Cells[1, 1, 1, $ws.Dimension.End.Column]
            $headerRange.Style.Font.Bold = $true
            $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $headerRange.Style.Fill.PatternType = 'Solid'
            $colorVerde = [System.Drawing.Color]::FromArgb(0, 176, 80)
            $headerRange.Style.Fill.BackgroundColor.SetColor($colorVerde)
        }
        Close-ExcelPackage $excelPackage
        
        Write-Host "-> Archivo '$nombreExcel' generado exitosamente." -ForegroundColor Green
    }
    catch {
        Write-Error "Ocurrió un error crítico al procesar '$nombreJson': $($_.Exception.Message)"
    }
}

Write-Host "`nConversión de todos los archivos finalizada."
