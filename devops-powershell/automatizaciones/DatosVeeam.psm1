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
    # Validacion: Verificar si Veeam fue seleccionado en el menu
    $veeamConn = $connections | Where-Object { $_.component -eq "veeam" }
    if (-not $veeamConn -or -not $veeamConn.conn) { return }

    Write-Host "-> Iniciando Recoleccion de Veeam..." -ForegroundColor Cyan

    # -------------------------------------------------------------------------
    # 1. RESOLUCION DE RUTAS (Metodo Seguro usando Config Global)
    # -------------------------------------------------------------------------
    # Usamos las variables del framework para no fallar con rutas relativas
    $bridgePath = Join-Path $global:CONFIG.PLUGINS_FOLDER "lib\VeeamBridge.ps1"
    $reportsDir = $global:CONFIG.REPORTS_FOLDER
    
    # El archivo resultado.txt esta en la raiz, un nivel arriba de devops.ps1
    # Calculamos la ruta absoluta para evitar errores de contexto
    $baseDir = (Get-Item $global:CONFIG.PLUGINS_FOLDER).Parent.Parent.FullName
    $resultTxtPath = Join-Path $baseDir "resultado.txt"

    # Convertimos rutas a absolutas para Start-Process
    $fullBridgePath = (Get-Item -Path $bridgePath -ErrorAction SilentlyContinue).FullName
    $fullReportsDir = (Get-Item -Path $reportsDir -ErrorAction SilentlyContinue).FullName

    # Validacion de seguridad
    if (-not $fullBridgePath -or -not (Test-Path $fullBridgePath)) {
        Write-Error "CRITICO: No se encuentra el script puente en: $bridgePath"
        return
    }

    # -------------------------------------------------------------------------
    # 2. CAPTURA INICIAL (Para detectar cambios)
    # -------------------------------------------------------------------------
    $existingFiles = @(Get-ChildItem -Path $fullReportsDir -Filter "VeeamBridge_*.json" -ErrorAction SilentlyContinue)

    # -------------------------------------------------------------------------
    # 3. EJECUCION DEL PUENTE (Tu Clase VeeamProactiva)
    # -------------------------------------------------------------------------
    # Intentamos usar PowerShell de 64-bits si estamos en un entorno de 32
    $psExe = "powershell.exe"
    $sysNative = "$env:windir\SysNative\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $sysNative) { $psExe = $sysNative }

    $processArgs = @(
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-NoExit", # Se mantiene abierta para ver errores (Tu clase la cierra con ENTER al final)
        "-File", "`"$fullBridgePath`""
    )

    try {
        Write-Host "   [LANZANDO] Abriendo consola nativa..." -ForegroundColor Cyan
        $p = Start-Process $psExe -ArgumentList $processArgs -PassThru
        $p.WaitForExit() # El script portable espera aqui
        Write-Host "   [FIN] Retornando al flujo principal." -ForegroundColor Green
    } catch {
        Write-Error "Fallo al iniciar puente: $($_.Exception.Message)"
        return
    }

    # -------------------------------------------------------------------------
    # 4. REGISTRO (Logica de DatosProactiva)
    # -------------------------------------------------------------------------
    # Buscamos archivos nuevos que no estaban en la captura inicial
    $currentFiles = @(Get-ChildItem -Path $fullReportsDir -Filter "VeeamBridge_*.json" -ErrorAction SilentlyContinue)
    
    $newFile = $currentFiles | Where-Object { 
        $name = $_.Name
        -not ($existingFiles.Name -contains $name)
    } | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($newFile) {
        Write-Host "   [DETECTADO] Nuevo reporte generado: $($newFile.Name)" -ForegroundColor Green
        
        try {
            # Escribimos SOLO el nombre del archivo en resultado.txt
            if (-not (Test-Path $resultTxtPath)) { New-Item $resultTxtPath -ItemType File -Force | Out-Null }
            
            Add-Content -Path $resultTxtPath -Value $newFile.Name -Force
            
            Write-Host "   [REGISTRADO] Agregado a resultado.txt correctamente." -ForegroundColor Green
        } catch {
            Write-Error "Error escribiendo en resultado.txt: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "No se genero ningun JSON nuevo. Revise la ventana azul por errores."
    }
}