using namespace System.Collections

# =============================================================================
# VEEAM BRIDGE - Ejecucion Nativa con Clase Estructurada
# =============================================================================
$ErrorActionPreference = "Stop"

Clear-Host
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "   CONSOLA NATIVA DE VEEAM (PUENTE DE EJECUCION)" -ForegroundColor Green
Write-Host "=========================================================="
Write-Host ""

class VeeamProactiva {
    # -------------------------------------------------------------------------
    # PROPIEDADES (Cada una sera una Hoja de Excel)
    # -------------------------------------------------------------------------
    [ArrayList]$Veeam_Repositories # Hoja 2: Repositorios
    [ArrayList]$Veeam_Proxies
    [ArrayList]$Veeam_ConfigDB
    [ArrayList]$Veeam_Configuration
    [ArrayList]$Veeam_Infrastructure
    
    # Estado interno
    [string]$ServerName
    [bool]$IsConnected

    # Constructor
    VeeamProactiva() {
        $this.Veeam_Repositories = [ArrayList]::new()
        $this.Veeam_Proxies = [ArrayList]::new()
        $this.Veeam_ConfigDB = [ArrayList]::new()
        $this.Veeam_Configuration = [ArrayList]::new()
        $this.Veeam_Infrastructure = [ArrayList]::new()
        $this.IsConnected = $false
        $this.ServerName = ""
    }

    # -------------------------------------------------------------------------
    # METODO 1: Carga de Modulos (Blindado)
    # -------------------------------------------------------------------------
    [void] CargarModulos() {
        Write-Host "[1/4] Buscando componentes de Veeam..." -ForegroundColor Cyan
        $loaded = $false
        
        # Rutas de busqueda fisica (Consola y Common)
        $paths = @(
            "C:\Program Files\Veeam\Backup and Replication\Console\Veeam.Backup.PowerShell",
            "$env:ProgramFiles\Veeam\Backup and Replication\Console\Veeam.Backup.PowerShell",
            "$env:ProgramFiles\Common Files\Veeam\PowerShell"
        )

        # 1. Intento por ruta absoluta .psd1
        foreach ($p in $paths) {
            if (Test-Path $p) {
                $m = Get-ChildItem -Path $p -Filter "Veeam.Backup.PowerShell.psd1" -Recurse -Depth 1 -ErrorAction SilentlyContinue | Select -First 1
                if ($m) {
                    try {
                        Import-Module $m.FullName -Scope Global -ErrorAction Stop
                        $loaded = $true; Write-Host " -> Modulo cargado por ruta." -ForegroundColor Gray; break
                    } catch {}
                }
            }
        }

        # 2. Intento Estandar
        if (-not $loaded -and (Get-Module -ListAvailable 'Veeam.Backup.PowerShell')) {
            try { Import-Module 'Veeam.Backup.PowerShell' -ErrorAction Stop; $loaded=$true } catch {}
        }

        # 3. Intento Snapin
        if (-not $loaded -and (Get-PSSnapin -Registered 'VeeamPSSnapIn')) {
            try { Add-PSSnapin 'VeeamPSSnapIn' -ErrorAction Stop; $loaded=$true } catch {}
        }

        if (-not $loaded) { throw "CRITICO: No se encontraron componentes de Veeam." }
    }

    # -------------------------------------------------------------------------
    # METODO 2: Conexion
    # -------------------------------------------------------------------------
    [void] Conectar() {
        Write-Host ""
        $s = Read-Host "Ingrese la IP o Hostname del Servidor Veeam"
        if ([string]::IsNullOrWhiteSpace($s)) { throw "Servidor invalido." }
        $this.ServerName = $s
        
        Write-Host "`n[2/4] Conectando a $s..."
        $c = Get-Credential -Message "Credenciales Veeam"
        
        try {
            $p = @{ Server=$s; Credential=$c; ErrorAction='Stop' }
            if (Get-Module 'Veeam.Backup.PowerShell') { $null = Veeam.Backup.PowerShell\Connect-VBRServer @p }
            else { $null = Connect-VBRServer @p }
            $this.IsConnected = $true
            Write-Host " CONEXION EXITOSA." -ForegroundColor Green
        } catch { throw "Error conectando: $($_.Exception.Message)" }
    }

    # -------------------------------------------------------------------------
    # METODO DE RECOLECCION 2: Repositorios (Genera Hoja 'Veeam_Repositories')
    # -------------------------------------------------------------------------
    [void] GetVeeamRepositories() {
        if (-not $this.IsConnected) { return }
        Write-Host " -> Recolectando Repositorios..." -ForegroundColor Gray

        try {
            $repos = Get-VBRBackupRepository
            foreach ($r in $repos) {
                $cap=0; $free=0; $perc=0
                try {
                    $cap = [math]::Round($r.Info.Capacity/1GB, 2)
                    $free = [math]::Round($r.Info.FreeSpace/1GB, 2)
                    if($cap -gt 0){ $perc = [math]::Round(($free/$cap)*100, 2) }
                } catch {}

                $row = [PSCustomObject]@{
                    "Repo Name"     = $r.Name
                    "Type"          = $r.Type.ToString()
                    "Path"          = $r.Path
                    "Capacity (GB)" = $cap
                    "Free (GB)"     = $free
                    "Free (%)"      = $perc
                }
                $this.Veeam_Repositories.Add($row) | Out-Null
            }
        } catch { Write-Warning "Error en Repos: $($_.Exception.Message)" }
    }

    
    [void] GetVeeamProxies() {
        if (-not $this.IsConnected) { return }
        Write-Host " -> Recolectando Proxies (VMware, Agentes, Archivos)..." -ForegroundColor Gray

        try {
            $processedIds = @{}
            $allProxies = [System.Collections.ArrayList]::new()
            
            # --- FUNCION LOCAL PARA AGREGAR SIN ERRORES DE TIPO ---
            $safeAdd = { param($list) 
                if ($list) {
                    foreach ($item in $list) { $allProxies.Add($item) | Out-Null }
                }
            }

            # 1. Computer File Proxy (El que fallaba antes)
            if (Get-Command "Get-VBRComputerFileProxyServer" -ErrorAction SilentlyContinue) {
                try { 
                    $p = Get-VBRComputerFileProxyServer -ErrorAction SilentlyContinue
                    & $safeAdd $p 
                } catch {}
            }

            # 2. VMware
            if (Get-Command "Get-VBRViProxy" -ErrorAction SilentlyContinue) {
                try { $p = Get-VBRViProxy -ErrorAction SilentlyContinue; & $safeAdd $p } catch {}
            }

            # 3. Hyper-V
            if (Get-Command "Get-VBRHvProxy" -ErrorAction SilentlyContinue) {
                try { $p = Get-VBRHvProxy -ErrorAction SilentlyContinue; & $safeAdd $p } catch {}
            }

            # 4. Agentes Backup
            if (Get-Command "Get-VBRComputerBackupProxy" -ErrorAction SilentlyContinue) {
                try { $p = Get-VBRComputerBackupProxy -ErrorAction SilentlyContinue; & $safeAdd $p } catch {}
            }

            # 5. NAS / File Share
            if (Get-Command "Get-VBRNasProxyServer" -ErrorAction SilentlyContinue) {
                try { $p = Get-VBRNasProxyServer -ErrorAction SilentlyContinue; & $safeAdd $p } catch {}
            }
            elseif (Get-Command "Get-VBRNasProxy" -ErrorAction SilentlyContinue) {
                try { $p = Get-VBRNasProxy -ErrorAction SilentlyContinue; & $safeAdd $p } catch {}
            }

            # 6. CDP
            if (Get-Command "Get-VBRCdpProxy" -ErrorAction SilentlyContinue) {
                try { $p = Get-VBRCdpProxy -ErrorAction SilentlyContinue; & $safeAdd $p } catch {}
            }
            
            # 7. Fallback General v12
            if (Get-Command "Get-VBRBackupProxy" -ErrorAction SilentlyContinue) {
                 try { $p = Get-VBRBackupProxy -ErrorAction SilentlyContinue; & $safeAdd $p } catch {}
            }

            # --- PROCESAMIENTO ---
            foreach ($proxy in $allProxies) {
                # ID Check
                $id = "NoID"
                if ($proxy.Id) { $id = $proxy.Id.ToString() }
                
                if ($id -ne "NoID" -and $processedIds.ContainsKey($id)) { continue }
                if ($id -ne "NoID") { $processedIds[$id] = $true }

                # EXTRACCION DE DATOS
                $proxyName = $proxy.Name 
                $type = "Standard"
                $maxTasks = 0

                # Deteccion de Tipo
                $className = $proxy.GetType().Name
                
                # CASO ESPECIAL: Agente de Archivo
                if ($className -match "ComputerFileProxyServer") {
                    $type = "Agent/File Proxy"
                    
                    # Nombre anidado
                    if ($proxy.Server -and $proxy.Server.Name) {
                        $proxyName = $proxy.Server.Name
                    }
                    # Tareas anidadas
                    if ($proxy.ConcurrentTaskNumber) {
                        $maxTasks = $proxy.ConcurrentTaskNumber
                    }
                }
                else {
                    # LOGICA ESTANDAR
                    if ($proxy.Type) { $type = $proxy.Type.ToString() }
                    elseif ($className -match "ViProxy") { $type = "VMware" }
                    elseif ($className -match "HvProxy") { $type = "Hyper-V" }
                    elseif ($className -match "ComputerBackupProxy") { $type = "Agent Backup" }
                    
                    if ($proxy.MaxTasksCount) { $maxTasks = $proxy.MaxTasksCount }
                }

                # Fallback nombre
                if ([string]::IsNullOrWhiteSpace($proxyName)) { $proxyName = "Unknown Proxy" }

                # HARDWARE INFO
                $hwData = $this.GetHardwareInfo($proxy, $proxyName)
                
                # CALCULOS
                $rec = $null; $over = $null
                if ($hwData.Cores -gt 0) {
                    $rec = 2 * $hwData.Cores
                    $over = $maxTasks -gt $rec
                }

                $this.AddProxyRow($proxyName, $type, $hwData.HostName, $hwData.Cores, $maxTasks, $rec, $over)
            }

        } catch { Write-Warning "Error procesando Proxies: $($_.Exception.Message)" }
    }

    # METODO: CONFIG DB & IMAGE (Hibrido: Registro + Fallbacks Indicativos)
    # -------------------------------------------------------------------------
    [void] GetVeeamConfigBackup() {
        if (-not $this.IsConnected) { return }
        Write-Host " -> Verificando proteccion de Base de Datos y Configuracion..." -ForegroundColor Gray

        try {
            # =========================================================
            # 1. BACKUP INTERNO DE CONFIGURACION (.bco)
            # =========================================================
            try {
                $configJob = Get-VBRConfigurationBackupJob -ErrorAction SilentlyContinue
                if ($configJob) {
                    $status = if ($configJob.Enabled) { "Enabled" } else { "Disabled" }
                    $lastResult = "Unknown"
                    $lastTime = "-"
                    
                    # Optimizacion: FindLastSession rapido
                    try {
                        $sess = $configJob.FindLastSession()
                        if ($sess) {
                            $lastResult = $sess.Result.ToString()
                            $lastTime = $sess.CreationTime.ToString("yyyy-MM-dd HH:mm")
                        } else { $lastResult = "Never Run" }
                    } catch { $lastResult = "Check Console" }

                    $this.Veeam_ConfigDB.Add([PSCustomObject]@{
                        "Check Type"     = "Veeam Config Backup (Native)"
                        "Target Name"    = "Veeam Configuration DB"
                        "Status"         = $status
                        "Last Result"    = $lastResult
                        "Last Run Time"  = $lastTime
                        "Details"        = "Respaldo .bco (Schedule: $($configJob.ScheduleOptions.BackupTime))"
                    }) | Out-Null
                }
            } catch { Write-Warning "Error leyendo ConfigJob: $($_.Exception.Message)" }

            # =========================================================
            # 2. DETECCION DE SQL SERVER (VIA REGISTRO)
            # =========================================================
            $dbServerName = $null
            $dbType = "Unknown"
            $regError = $false

            try {
                $baseKey = 'HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication\DatabaseConfigurations'
                
                # Verificamos si existe la ruta (Si corre en consola remota, esto dara falso)
                if (Test-Path $baseKey) {
                    $active = (Get-ItemProperty -Path $baseKey -Name 'SqlActiveConfiguration' -ErrorAction SilentlyContinue).SqlActiveConfiguration
                    
                    if ($active -eq "MsSql") {
                        $dbType = "MS SQL"
                        $dbServerName = (Get-ItemProperty -Path "$baseKey\MsSql" -Name 'SqlServerName' -ErrorAction SilentlyContinue).SqlServerName
                    }
                    elseif ($active -eq "PostgreSql") {
                        $dbType = "PostgreSQL"
                        $dbServerName = (Get-ItemProperty -Path "$baseKey\PostgreSql" -Name 'SqlHostName' -ErrorAction SilentlyContinue).SqlHostName
                    }
                } else {
                    $regError = $true # No existe la ruta (ej: consola remota)
                }
            } catch {
                $regError = $true # Error de permisos
            }

            # Si fallo el registro, llenamos de forma indicativa
            if (-not $dbServerName) {
                $msg = if ($regError) { "No accesible (Consola Remota/Permisos)" } else { "Unknown" }
                
                $this.Veeam_ConfigDB.Add([PSCustomObject]@{
                    "Check Type"     = "DB Connection Info"
                    "Target Name"    = "SQL/Postgres Server"
                    "Status"         = "INFO"
                    "Last Result"    = "-"
                    "Last Run Time"  = "-"
                    "Details"        = $msg
                }) | Out-Null
                
                # FALLBACK: Usamos el nombre del servidor Veeam conectado
                $dbServerName = $this.ServerName
            } else {
                 # Normalizar localhost
                 if ($dbServerName -in @("localhost", "(local)", "127.0.0.1", ".")) {
                    $dbServerName = $env:COMPUTERNAME
                 }
            }

            # =========================================================
            # 3. BUSQUEDA DE BACKUP DE IMAGEN (VM/Agente)
            # =========================================================
            # Buscamos backups para el servidor detectado (SQL o el propio Veeam si fallo registro)
            
            # Limpieza nombre
            if ($dbServerName -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}") {
                try { $dns=[System.Net.Dns]::GetHostEntry($dbServerName); $dbServerName=$dns.HostName } catch {}
            }

            $foundPoints = $null
            # Intento A: Exacto
            $foundPoints = Get-VBRRestorePoint -Name $dbServerName -ErrorAction SilentlyContinue
            # Intento B: Aproximado
            if (-not $foundPoints) {
                $short = $dbServerName.Split('.')[0]
                $foundPoints = Get-VBRRestorePoint | Where-Object { $_.VMName -eq $short -or $_.VMName -like "$short.*" }
            }

            if ($foundPoints) {
                $lastRp = $foundPoints | Sort-Object CreationTime -Descending | Select-Object -First 1
                $jName = "Unknown Job"
                if ($lastRp.JobName) { $jName = $lastRp.JobName }
                elseif ($lastRp.Backup) { $jName = $lastRp.Backup.JobName }

                $this.Veeam_ConfigDB.Add([PSCustomObject]@{
                    "Check Type"     = "Server/DB Image Backup"
                    "Target Name"    = $dbServerName
                    "Status"         = "Protected"
                    "Last Result"    = "OK"
                    "Last Run Time"  = $lastRp.CreationTime.ToString("yyyy-MM-dd HH:mm")
                    "Details"        = "Backup encontrado (Job: $jName)"
                }) | Out-Null
            } else {
                $this.Veeam_ConfigDB.Add([PSCustomObject]@{
                    "Check Type"     = "Server/DB Image Backup"
                    "Target Name"    = $dbServerName
                    "Status"         = "WARNING - Not Found"
                    "Last Result"    = "-"
                    "Last Run Time"  = "-"
                    "Details"        = "No se encontro backup de imagen para este servidor."
                }) | Out-Null
            }

        } catch { Write-Warning "Error verificando Config Backup: $($_.Exception.Message)" }
    }

    [void] GetVeeamConfigurationCheck() {
        if (-not $this.IsConnected) { return }
        Write-Host " -> Verificando Config Backup (.NET Core)..." -ForegroundColor Gray

        try {
            $config = Get-VBRConfigurationBackupJob -ErrorAction SilentlyContinue
            
            if ($config) {
                
                # --- TRUCO: Ejecucion Dinamica para evitar Error de Parseo ---
                # Definimos el comando como texto para que PowerShell no intente 
                # validar el tipo [Veeam...] antes de cargar los modulos.
                
                $cmdCore = "[Veeam.Backup.Core.CConfigurationBackupJob]::Find()"
                $configJobCore = Invoke-Expression $cmdCore

                # Ahora usamos el objeto que recuperamos
                $configLastSession = $configJobCore.FindLastSessionByResults('Success')

                $lastDate = "Never"
                if ($configLastSession) {
                    $lastDate = $configLastSession.CreationTime.ToString("yyyy-MM-dd HH:mm")
                }

                $row = [PSCustomObject]@{
                    Enabled        = $config.Enabled
                    Repository     = if ($config.Repository) { $config.Repository.Name } else { "Unknown" }
                    Schedule       = $config.ScheduleOptions.Enabled
                    TipoDeSchedule = $config.ScheduleOptions.Type.ToString()
                    UltimoBackup   = $lastDate
                }

                $this.Veeam_Configuration.Add($row) | Out-Null
            }
        } catch { Write-Warning "Error en ConfigBackup: $($_.Exception.Message)" }
    }

    [void] GetVeeamInfrastructureAlerts() {
        if (-not $this.IsConnected) { return }
        Write-Host " -> Verificando Paquetes Pendientes y Versiones..." -ForegroundColor Gray

        try {
            # 1. VERSION MAESTRA
            $masterVersion = "0.0.0.0"
            try {
                $path = "$env:ProgramFiles\Veeam\Backup and Replication\Console\Veeam.Backup.Shell.exe"
                if (Test-Path $path) {
                    $masterVersion = (Get-Item $path).VersionInfo.ProductVersion
                }
            } catch {}
            Write-Host "    > Version Maestra: $masterVersion" -ForegroundColor Cyan

            # 2. ANALISIS
            $servers = Get-VBRServer
            
            foreach ($s in $servers) {
                $issues = [System.Collections.ArrayList]::new()
                $debugMode = ($s.Name -like "*ArBAVSPBKP01*") # TU SERVER PROBLEMATICO
                
                if ($debugMode) { Write-Host "    > [ESPIA] Analizando $s.Name..." -ForegroundColor Magenta }

                # A. Disponibilidad
                if ($s.IsUnavailable) { $issues.Add("Unavailable") | Out-Null }

                # B. Reinicio
                try {
                    if ($s.GetPhysicalHost().IsRebootRequired()) { $issues.Add("Reboot Required") | Out-Null }
                } catch {}

                # C. PAQUETES A DESPLEGAR (La prueba definitiva de Missing Updates)
                try {
                    $phys = $s.GetPhysicalHost()
                    if ($phys.PSObject.Methods["GetPackagesToDeploy"]) {
                        $pkgs = $phys.GetPackagesToDeploy()
                        if ($pkgs -and $pkgs.Count -gt 0) {
                            $issues.Add("Missing Packages ($($pkgs.Count))") | Out-Null
                            if ($debugMode) { 
                                Write-Host "      * Paquetes pendientes encontrados!" -ForegroundColor Yellow
                                foreach ($p in $pkgs) { Write-Host "        - Pkg: $($p.Type) / Path: $($p.DistributivePath)" -ForegroundColor DarkGray }
                            }
                        }
                    }
                } catch { if($debugMode){Write-Host "      * Error leyendo GetPackagesToDeploy: $($_.Exception.Message)" -ForegroundColor Red} }

                # D. COMPARACION VERSIONES (Verbose para tu server)
                try {
                    $phys = $s.GetPhysicalHost()
                    if ($phys.PSObject.Methods["GetInstalledComponents"]) {
                        $comps = $phys.GetInstalledComponents()
                        
                        foreach ($c in $comps) {
                            # Imprimir todo lo que tiene tu server
                            if ($debugMode) {
                                Write-Host "      * Instalado: $($c.Type) = $($c.Version)" -ForegroundColor Gray
                            }

                            if ($c.Version -and $c.Version -ne $masterVersion) {
                                $vLocal  = [version]$c.Version
                                $vMaster = [version]$masterVersion
                                
                                if ($vLocal -lt $vMaster) {
                                    $issues.Add("Out of Date ($($c.Type))") | Out-Null
                                }
                            }
                        }
                    } else {
                        if ($debugMode) { Write-Host "      * No tiene metodo GetInstalledComponents" -ForegroundColor Red }
                    }
                } catch {}

                # E. CHECK BOOLEANO (Ultimo recurso)
                if ($issues.Count -eq 0) {
                    try {
                        if ($s.GetPhysicalHost().IsComponentsUpdateRequired()) {
                            $issues.Add("Out of Date (Flag)") | Out-Null
                        }
                    } catch {}
                }

                # REPORTE
                $uniqIssues = $issues | Select-Object -Unique
                if ($uniqIssues) {
                    $issueStr = $uniqIssues -join ", "
                    Write-Host "    [ALERTA] $($s.Name): $issueStr" -ForegroundColor Yellow
                    
                    $this.Veeam_Infrastructure.Add([PSCustomObject]@{
                        "Component Name" = $s.Name
                        "Type"           = $s.Type.ToString()
                        "Description"    = $s.Description
                        "Status"         = "ALERT"
                        "Details"        = $issueStr
                    }) | Out-Null
                } elseif ($debugMode) {
                    Write-Host "      * CONCLUSION ESPIA: El script no detecto problemas (Issues count = 0)." -ForegroundColor Green
                }
            }

        } catch { Write-Warning "Error en Infraestructura: $($_.Exception.Message)" }
    }

    # --- FUNCION AUXILIAR (V3: BUSQUEDA IMPLACABLE) ---
    [pscustomobject] GetHardwareInfo($obj, $name) {
        $res = [PSCustomObject]@{ Cores = 0; HostName = $name }
        
        # --- INTENTO 1: PROPIEDADES DIRECTAS (VMware/Hyper-V) ---
        try {
            if ($obj.Host -and $obj.Host.HardwareInfo) { 
                $res.Cores = $obj.Host.HardwareInfo.CoresCount
                $res.HostName = $obj.Host.Name
                return $res
            }
            if ($obj.Host -and $obj.Host.GetPhysicalHost) { 
                $phys = $obj.Host.GetPhysicalHost()
                if ($phys.HardwareInfo) {
                    $res.Cores = $phys.HardwareInfo.CoresCount
                    $res.HostName = $obj.Host.Name
                    return $res
                }
            }
        } catch {}

        # --- INTENTO 2: BUSQUEDA EN INFRAESTRUCTURA (Get-VBRServer) ---
        # Traemos TODOS los servidores para evitar errores de "Exact Match" en el nombre
        try {
            $allServers = Get-VBRServer -ErrorAction SilentlyContinue
            $targetServer = $null

            # A. Busqueda Exacta
            $targetServer = $allServers | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            
            # B. Busqueda por Nombre Corto (Si falla la exacta)
            if (-not $targetServer) {
                $shortName = $name.Split('.')[0]
                $targetServer = $allServers | Where-Object { $_.Name.Split('.')[0] -eq $shortName } | Select-Object -First 1
            }

            if ($targetServer) {
                $res.HostName = $targetServer.Name # Actualizamos con el nombre real registrado
                
                # Leemos Hardware desde el objeto VBRServer encontrado
                if ($targetServer.Info -and $targetServer.Info.HardwareInfo) {
                     $res.Cores = $targetServer.Info.HardwareInfo.CoresCount
                     return $res
                }
                if ($targetServer.HardwareInfo) {
                     $res.Cores = $targetServer.HardwareInfo.CoresCount
                     return $res
                }
                if ($targetServer.Model -and $targetServer.Model.HardwareInfo) {
                     $res.Cores = $targetServer.Model.HardwareInfo.CoresCount
                     return $res
                }
                # Metodo explicito
                if ($targetServer.PSObject.Methods["GetHardwareInfo"]) {
                    try {
                        $hw = $targetServer.GetHardwareInfo()
                        if ($hw) { $res.Cores = $hw.CoresCount; return $res }
                    } catch {}
                }
            }
        } catch {}

        # --- INTENTO 3: WMI DIRECTO (ULTIMO RECURSO) ---
        # Si Veeam no sabe, le preguntamos al Windows directamente.
        # Esto funciona si el usuario que corre el script tiene acceso admin al proxy.
        if ($res.Cores -eq 0) {
            try {
                # Intentamos conectar por WMI al servidor remoto
                $wmi = Get-WmiObject -Class Win32_Processor -ComputerName $name -ErrorAction SilentlyContinue
                if ($wmi) {
                    # Sumamos los cores de todos los sockets
                    $totalCores = ($wmi | Measure-Object -Property NumberOfCores -Sum).Sum
                    if ($totalCores -gt 0) {
                        $res.Cores = $totalCores
                        return $res
                    }
                }
            } catch {}
        }
        
        return $res
    }

    # --- FUNCION AUXILIAR PARA AGREGAR FILA (CORREGIDO: $realHost en lugar de $host) ---
    [void] AddProxyRow($name, $type, $realHost, $cores, $max, $rec, $warn) {
        $row = [PSCustomObject]@{
            "Name"           = $name
            "Type"           = $type
            "Host Real"      = $realHost
            "CpuCores"       = if($cores -gt 0){$cores}else{"N/A"}
            "MaxTasks"       = $max
            "RecommendedMax" = if($rec -gt 0){$rec}else{"-"}
            "Warning"        = $warn
        }
        $this.Veeam_Proxies.Add($row) | Out-Null
    }

    [hashtable] GetReporteFinal() {
        # Retornamos un Hashtable donde:
        # Key (Nombre) = Nombre de la Hoja en Excel
        # Value (Datos) = Array de filas
        return @{
            "Veeam Jobs"         = $this.Veeam_Jobs
            "Veeam Repositories" = $this.Veeam_Repositories
            "Proxies"            = $this.Veeam_Proxies
            "Config DB"          = $this.Veeam_ConfigDB
            "Veeam BK Config"    = $this.Veeam_Configuration
            "Components Status"  = $this.Veeam_Infrastructure
        }
    }

    [void] GuardarJSON() {
        Write-Host "`n[4/4] Guardando resultados..."
        try {
            # Calculo de ruta seguro
            $base = $PSScriptRoot; if(!$base){$base=(Get-Location).Path}
            $repDir = Join-Path (Split-Path (Split-Path $base -Parent) -Parent) "reportes"
            if(!(Test-Path $repDir)){ $repDir="C:\Temp\WPC_Reportes"; New-Item -Type Directory -Force $repDir|Out-Null }

            $file = "VeeamBridge_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $path = Join-Path $repDir $file

            # Estructura compatible con JSONtoExcels
            $data = $this.GetReporteFinal()
            
            $payload = [PSCustomObject]@{
                Result    = "OK"
                Name      = "VeeamProactiva" # ID del reporte
                DateTime  = (Get-Date -Format "yyyy-MM-dd HH:mm")
                User      = $env:USERNAME
                Endpoint  = $this.ServerName
                Component = "veeam"
                Report    = $data  # Aqui va el Hashtable con las hojas
            }

            $payload | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
            Write-Host " [OK] Archivo: $path" -ForegroundColor Green
        } catch { Write-Error "Fallo guardando JSON: $($_.Exception.Message)" }
    }
}

# =============================================================================
# BLOQUE DE EJECUCION PRINCIPAL
# =============================================================================
try {
    # 1. Instanciar
    $proactiva = [VeeamProactiva]::new()

    # 2. Preparar
    $proactiva.CargarModulos()
    $proactiva.Conectar()

    # 3. Ejecutar Metodos de Recoleccion (Cada uno llena su lista)
    Write-Host "`n[3/4] Ejecutando tareas de recoleccion..." -ForegroundColor Cyan
    $proactiva.GetVeeamRepositories()
    $proactiva.GetVeeamProxies()
    $proactiva.GetVeeamConfigBackup()
    $proactiva.GetVeeamConfigurationCheck()
    $proactiva.GetVeeamInfrastructureAlerts()

    # 4. Finalizar
    $proactiva.GuardarJSON()

} catch {
    Write-Host "`nERROR FATAL: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nPuede cerrar esta ventana."
Read-Host "Presione ENTER..."