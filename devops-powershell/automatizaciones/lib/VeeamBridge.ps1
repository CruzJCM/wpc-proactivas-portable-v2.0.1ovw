# =============================================================================
# VEEAM BRIDGE - Ejecucion en Entorno Nativo
# =============================================================================
$ErrorActionPreference = "Stop"

# Limpieza inicial
Clear-Host
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "   CONSOLA NATIVA DE VEEAM (PUENTE DE EJECUCION)" -ForegroundColor Green
Write-Host "=========================================================="
Write-Host ""

# 1. CARGA DE MODULOS (Estrategia de Busqueda Fisica Blindada)
Write-Host "[1/3] Buscando componentes de Veeam en el disco..." -ForegroundColor Cyan

$veeamModuleLoaded = $false
$moduleName = ""

# Lista de rutas donde buscar el modulo (Prioridad: Consola > Common)
# Buscamos el archivo real .psd1 para cargarlo por ruta absoluta
$searchPaths = @(
    "C:\Program Files\Veeam\Backup and Replication\Console\Veeam.Backup.PowerShell",
    "$env:ProgramFiles\Veeam\Backup and Replication\Console\Veeam.Backup.PowerShell",
    "$env:ProgramFiles\Common Files\Veeam\PowerShell"
)

# Intentamos encontrar el archivo .psd1 real
foreach ($path in $searchPaths) {
    if (Test-Path $path) {
        $manifest = Get-ChildItem -Path $path -Filter "Veeam.Backup.PowerShell.psd1" -Recurse -Depth 1 -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if ($manifest) {
            Write-Host " -> Encontrado en: $($manifest.FullName)" -ForegroundColor Gray
            try {
                # Importamos usando la ruta completa. 
                # Al estar en powershell.exe nativo, esto SI funcionara y vera las DLLs vecinas.
                Import-Module -Name $manifest.FullName -Scope Global -ErrorAction Stop
                $veeamModuleLoaded = $true
                $moduleName = 'Veeam.Backup.PowerShell'
                Write-Host " [OK] Modulo cargado por ruta absoluta." -ForegroundColor Green
                break
            } catch {
                Write-Warning " Error al importar desde ruta: $($_.Exception.Message)"
            }
        }
    }
}

# Si fallo la busqueda fisica, intentamos metodo estandar (fallback)
if (-not $veeamModuleLoaded) {
    if (Get-Module -ListAvailable -Name 'Veeam.Backup.PowerShell') {
        try {
            Import-Module 'Veeam.Backup.PowerShell' -ErrorAction Stop
            Write-Host ' [OK] Modulo cargado (Metodo Estandar).' -ForegroundColor Green
            $veeamModuleLoaded = $true
            $moduleName = 'Veeam.Backup.PowerShell'
        } catch {}
    }
}

# Ultimo recurso: Snapin Legacy
if (-not $veeamModuleLoaded) {
    if (Get-PSSnapin -Registered -Name VeeamPSSnapIn -ErrorAction SilentlyContinue) {
        try {
            Add-PSSnapin 'VeeamPSSnapin' -ErrorAction Stop
            Write-Host ' [OK] Snapin Legacy cargado.' -ForegroundColor Green
            $veeamModuleLoaded = $true
            $moduleName = 'VeeamPSSnapin'
        } catch {}
    }
}

# Verificacion final
if (-not $veeamModuleLoaded) {
    Write-Host ""
    Write-Host "ERROR CRITICO: No se encontraron los componentes de PowerShell de Veeam." -ForegroundColor Red
    Write-Host "Diagnostico: Se busco en rutas estandar y en el entorno actual." -ForegroundColor Gray
    Write-Host "La ventana heredada puede tener las rutas bloqueadas."
    Read-Host "Presione ENTER para salir..."
    exit
}

# 2. CONEXION
Write-Host ""
$veeamServer = Read-Host "Ingrese la IP o Hostname del Servidor Veeam"
if ([string]::IsNullOrWhiteSpace($veeamServer)) { exit }

try {
    $cred = Get-Credential -Message "Credenciales para Veeam en $veeamServer"
    
    Write-Host "`n[2/3] Conectando a $veeamServer..."
    
    $connParams = @{
        Server      = $veeamServer
        Credential  = $cred
        ErrorAction = 'Stop'
    }

    if ($moduleName -eq 'Veeam.Backup.PowerShell') {
        $session = Veeam.Backup.PowerShell\Connect-VBRServer @connParams
    } else {
        $session = Connect-VBRServer @connParams
    }
    
    Write-Host " CONEXION EXITOSA." -ForegroundColor Green

} catch {
    Write-Host " FALLO LA CONEXION." -ForegroundColor Red
    Write-Warning "Detalle: $($_.Exception.Message)"
    Read-Host "Presione ENTER para salir..."
    exit
}

# 3. RECOLECCION DE DATOS
Write-Host "`n[3/3] Recolectando informacion..." 
$finalReport = @()

# --- A. JOBS ---
try {
    Write-Host " -> Procesando Jobs..." -ForegroundColor Gray
    $jobs = Get-VBRJob | Where-Object { $_.IsScheduleEnabled -eq $true }
    
    foreach ($job in $jobs) {
        $last = $job.FindLastSession()
        $finalReport += [PSCustomObject]@{
            Type          = "Job"
            Name          = $job.Name
            Status        = if ($last) { $last.Result.ToString() } else { "Never Run" }
            LastRun       = if ($last) { $last.CreationTime.ToString("yyyy-MM-dd HH:mm") } else { "-" }
            NextRun       = $job.Info.NextRunTime.ToString("yyyy-MM-dd HH:mm")
            CapacityGB    = 0
            FreeGB        = 0
            FreePercent   = 0
        }
    }
} catch { Write-Warning "Error leyendo Jobs: $($_.Exception.Message)" }

# --- B. REPOSITORIOS ---
try {
    Write-Host " -> Procesando Repositorios..." -ForegroundColor Gray
    $repos = Get-VBRBackupRepository
    
    foreach ($repo in $repos) {
        $cap = 0; $free = 0; $perc = 0
        try {
            $cap = [math]::Round($repo.Info.Capacity / 1GB, 2)
            $free = [math]::Round($repo.Info.FreeSpace / 1GB, 2)
            if ($cap -gt 0) { $perc = [math]::Round(($free / $cap) * 100, 2) }
        } catch {}

        $finalReport += [PSCustomObject]@{
            Type          = "Repo"
            Name          = $repo.Name
            Status        = "OK"
            LastRun       = "-"
            NextRun       = "-"
            CapacityGB    = $cap
            FreeGB        = $free
            FreePercent   = $perc
        }
    }
} catch { Write-Warning "Error leyendo Repositorios: $($_.Exception.Message)" }

# 4. EXPORTACION (JSON BRIDGE)
try {
    # Truco para encontrar la carpeta reportes subiendo niveles desde este script
    # Asumimos estructura: devops-powershell\automatizaciones\lib\VeeamBridge.ps1
    $scriptPath = $MyInvocation.MyCommand.Path
    # Subimos 3 niveles: lib -> automatizaciones -> devops-powershell -> [Raiz con reportes]
    # O ajustamos segun tu estructura: devops-powershell\reportes
    
    # Buscamos la carpeta reportes relativa al script
    $libDir = Split-Path $scriptPath -Parent
    $autoDir = Split-Path $libDir -Parent
    $devopsDir = Split-Path $autoDir -Parent
    $reportsDir = Join-Path $devopsDir "reportes"
    
    if (-not (Test-Path $reportsDir)) { 
        # Fallback seguro a C:\Temp si la estructura de carpetas es rara o falla
        $reportsDir = "C:\Temp\WPC_Reportes"
        New-Item -ItemType Directory -Force -Path $reportsDir | Out-Null 
    }

    $fileName = "VeeamBridge_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $exportPath = Join-Path $reportsDir $fileName

    $payload = [PSCustomObject]@{
        Result    = "OK"
        Name      = "VeeamBridgeData"
        DateTime  = (Get-Date -Format "yyyy-MM-dd HH:mm")
        User      = $env:USERNAME
        Endpoint  = $veeamServer
        Component = "veeam"
        Report    = $finalReport
    }

    $payload | ConvertTo-Json -Depth 5 | Set-Content -Path $exportPath -Encoding UTF8
    
    Write-Host ""
    Write-Host " [OK] Reporte generado exitosamente:" -ForegroundColor Green
    Write-Host " $exportPath" -ForegroundColor White
} catch {
    Write-Error "Error guardando el JSON: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Puede cerrar esta ventana (Presione ENTER)."
Read-Host