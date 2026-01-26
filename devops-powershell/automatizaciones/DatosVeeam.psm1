# ==============================================================================
# PLUGIN: DatosVeeam.psm1
# Ubicación: devops-powershell/automatizaciones/DatosVeeam.psm1
# ==============================================================================

# Importamos la librería que creamos en el Paso 2
#using module .\lib\VeeamBridge.ps1

function Get-DatosVeeam($connections) {
    # Función de entrada que busca el framework.
    # El nombre debe coincidir con el del archivo (menos la extensión) o el patrón que use tu lanzador.
    Start-DatosVeeam($connections)

    <#
    .Synopsis
    Recolecta datos básicos de Veeam para prueba de conectividad.
    .Component
    veeam
    .Role
    ui
    #>
}


function Start-DatosVeeam($connections) {
    # Validación básica de que Veeam está activo
    $veeamConn = $connections | Where-Object { $_.component -eq "veeam" }
    if (-not $veeamConn -or -not $veeamConn.conn) { return }

    Write-Host "-> Iniciando Recolección de Veeam..." -ForegroundColor Cyan

    # -------------------------------------------------------------------------
    # 1. RESOLUCIÓN DE RUTAS (MÉTODO SEGURO)
    # -------------------------------------------------------------------------
    # Usamos las variables de configuración del framework para no fallar
    
    # Ruta del script puente (dentro de la carpeta de plugins/lib)
    $bridgePath = Join-Path $global:CONFIG.PLUGINS_FOLDER "lib\VeeamBridge.ps1"
    
    # Ruta de reportes
    $reportsDir = $global:CONFIG.REPORTS_FOLDER
    
    # Ruta de resultado.txt (Asumimos que está un nivel arriba de donde corre el script base)
    # devops.ps1 corre en /devops-powershell/, así que resultado.txt está en ../
    $resultTxtPath = "..\resultado.txt"

    # Convertimos a rutas absolutas para que Start-Process no se confunda
    $fullBridgePath = (Get-Item -Path $bridgePath -ErrorAction SilentlyContinue).FullName
    $fullReportsDir = (Get-Item -Path $reportsDir -ErrorAction SilentlyContinue).FullName

    # --- VALIDACIÓN DE SEGURIDAD ---
    if (-not $fullBridgePath -or -not (Test-Path $fullBridgePath)) {
        Write-Error "CRÍTICO: No se encuentra el script puente en: $bridgePath"
        return # Aquí es donde probablemente fallaba antes
    }

    # -------------------------------------------------------------------------
    # 2. CAPTURA DE ESTADO INICIAL
    # -------------------------------------------------------------------------
    $existingFiles = @(Get-ChildItem -Path $fullReportsDir -Filter "VeeamBridge_*.json" -ErrorAction SilentlyContinue)

    # -------------------------------------------------------------------------
    # 3. LANZAR PUENTE (Ventana Azul)
    # -------------------------------------------------------------------------
    $psExe = "powershell.exe"
    $sysNative = "$env:windir\SysNative\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $sysNative) { $psExe = $sysNative }

    $processArgs = @(
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-NoExit", # Mantener abierto para debug (el script cierra con ENTER)
        "-File", "`"$fullBridgePath`""
    )

    try {
        Write-Host "   [LANZANDO] Abriendo consola nativa..." -ForegroundColor Cyan
        $p = Start-Process $psExe -ArgumentList $processArgs -PassThru
        $p.WaitForExit() 
        Write-Host "   [FIN] Retornando al flujo principal." -ForegroundColor Green
    } catch {
        Write-Error "Fallo al iniciar puente: $($_.Exception.Message)"
        return
    }

    # -------------------------------------------------------------------------
    # 4. REGISTRAR RESULTADO
    # -------------------------------------------------------------------------
    # Buscamos archivos nuevos
    $currentFiles = @(Get-ChildItem -Path $fullReportsDir -Filter "VeeamBridge_*.json" -ErrorAction SilentlyContinue)
    
    $newFile = $currentFiles | Where-Object { 
        $name = $_.Name
        -not ($existingFiles.Name -contains $name)
    } | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($newFile) {
        Write-Host "   [DETECTADO] Reporte generado: $($newFile.Name)" -ForegroundColor Green
        
        try {
            # Escribir en resultado.txt
            if (-not (Test-Path $resultTxtPath)) { New-Item $resultTxtPath -ItemType File -Force | Out-Null }
            Add-Content -Path $resultTxtPath -Value $newFile.Name -Force
            Write-Host "   [REGISTRADO] Agregado a resultado.txt correctamente." -ForegroundColor Green
        } catch {
            Write-Error "Error escribiendo resultado.txt: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "No se generó ningún JSON nuevo."
    }
}