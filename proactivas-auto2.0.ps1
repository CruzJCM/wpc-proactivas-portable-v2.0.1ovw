#Author: Santiago Alvarez
#Githb: SanntAlv

if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Force -AllowClobber
}

$baseDir = $PSScriptRoot 

if (-not $path){
    $path = (Get-Item -Path ".\" -Verbose).FullName
}

$nombreCliente = Read-Host "Escriba el Nombre del cliente"
$mes = Read-Host "Escriba el mes correspondiente a la tarea Proactiva (ej: Enero, Febrero...)"

$rutaArchivos = Join-Path -Path $baseDir -ChildPath "devops-powershell\reportes\proactiva-excel"    
$archivoSalida = "${baseDir}\devops-powershell\reportes\Anexo\Anexo Tecnico - ${nombreCliente} - ${mes}.xlsx"
$rutaSalidaChecklist = "${baseDir}\devops-powershell\reportes\Anexo\Checklist - ${nombreCliente} - ${mes}.xlsx"
$rutaPlantilla = "${baseDir}\devops-powershell\reportes\templete\Checklist Proactiva Actualizada - Cliente - Mes.xlsx"


$excelMasReciente = Get-ChildItem -Path $rutaArchivos -Filter "*_Proactiva*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

$datosSalida = @()

function Exportar-InformeConEstilo {
    param(
        [Parameter(Mandatory=$true)]
        [object[]]$Datos, 

        [Parameter(Mandatory=$true)]
        [string]$RutaArchivo, 

        [Parameter(Mandatory=$true)]
        [string]$NombreHoja, 

        [string[]]$ColumnasSinConversion 
    )
    
    $exportParams = @{
        Path = $RutaArchivo
        WorksheetName = $NombreHoja
        AutoSize = $true
        FreezeTopRow = $true
        PassThru = $true 
    }

    if ($PSBoundParameters.ContainsKey('ColumnasSinConversion')) {
        $exportParams['NoNumberConversion'] = $ColumnasSinConversion
    }
    
    $excelPackage = $Datos | Export-Excel @exportParams

    $nombreHojaReal = $NombreHoja.Substring(0, [Math]::Min(31, $NombreHoja.Length))
    
    $ws = $excelPackage.Workbook.Worksheets[$nombreHojaReal]
    
    if ($null -ne $ws -and $null -ne $ws.Dimension) {
        $headerRange = $ws.Cells[1, 1, 1, $ws.Dimension.End.Column]
        
        $headerRange.Style.Font.Bold = $true
        $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $headerRange.Style.Fill.PatternType = 'Solid'
        
        $colorVerde = [System.Drawing.Color]::FromArgb(0, 176, 80)
        $headerRange.Style.Fill.BackgroundColor.SetColor($colorVerde)
    } else {
        Write-Warning "No se pudo aplicar estilo a la hoja '$nombreHojaReal' porque no se encontro o estaba vacía."
    }

    $excelPackage.Save()
    $excelPackage.Dispose()
}

function AnalizarSize {

    $sizingWeights = @{
        'tiny'    = 1; 'small'   = 2; 'medium'  = 3
        'large'   = 4; 'x-large' = 5
    }

    $sizingByVersion = @{ 
        '9' = @{    #igual que version 8
            requisitos = @{
                'tiny'    = @{ Cores = 2;  MemoriaGB = 14; VMsMax = 100 }
                'small'   = @{ Cores = 4;  MemoriaGB = 21; VMsMax = 1000 }
                'medium'  = @{ Cores = 8;  MemoriaGB = 30; VMsMax = 4000 }
                'large'   = @{ Cores = 16; MemoriaGB = 39; VMsMax = 10000 }
                'x-large' = @{ Cores = 24; MemoriaGB = 58; VMsMax = 35000 }
            }
        }
        '8' = @{ 
            requisitos = @{
                'tiny'    = @{ Cores = 2;  MemoriaGB = 14; VMsMax = 100 }
                'small'   = @{ Cores = 4;  MemoriaGB = 21; VMsMax = 1000 }
                'medium'  = @{ Cores = 8;  MemoriaGB = 30; VMsMax = 4000 }
                'large'   = @{ Cores = 16; MemoriaGB = 39; VMsMax = 10000 }
                'x-large' = @{ Cores = 24; MemoriaGB = 58; VMsMax = 35000 }
            }
        }
        '7' = @{
            requisitos = @{
                'tiny'    = @{ Cores = 2;  MemoriaGB = 12; VMsMax = 100 }
                'small'   = @{ Cores = 4;  MemoriaGB = 19; VMsMax = 1000 }
                'medium'  = @{ Cores = 8;  MemoriaGB = 28; VMsMax = 4000 }
                'large'   = @{ Cores = 16; MemoriaGB = 37; VMsMax = 10000 }
                'x-large' = @{ Cores = 24; MemoriaGB = 56; VMsMax = 35000 }
            }
        }
        '6' = @{ #6.7 = 6.5
            requisitos = @{
                'tiny'    = @{ Cores = 2;  MemoriaGB = 10; VMsMax = 100 }
                'small'   = @{ Cores = 4;  MemoriaGB = 16; VMsMax = 1000 }
                'medium'  = @{ Cores = 8;  MemoriaGB = 24; VMsMax = 4000 }
                'large'   = @{ Cores = 16; MemoriaGB = 32; VMsMax = 10000 }
                'x-large' = @{ Cores = 24; MemoriaGB = 48; VMsMax = 35000 }
            }
        }
    }

    $mapaDeEstadosVM = @{}; $datosVCenter = @(); $datosSizing = @()
    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivo = $_.FullName
        (Import-Excel -Path $archivo -WorksheetName "VM") | ForEach-Object {
            if (-not $mapaDeEstadosVM.ContainsKey($_.VM)) { $mapaDeEstadosVM.Add($_.VM, $_.State) }
        }
        $datosVCenter += Import-Excel -Path $archivo -WorksheetName "vCenter"
        $datosSizing += Import-Excel -Path $archivo -WorksheetName "Sizing"
    }

    $datosCombinados = @()

    foreach ($fila in $datosSizing) {
        $estado = $mapaDeEstadosVM[$fila.VM]
        if ($estado -ne "PoweredOff") {
            $infoVC = $datosVCenter | Where-Object { $_.'vCenter Server' -eq $fila.vCenter } | Select-Object -First 1
            $datosCombinados += [PSCustomObject]@{
                State = $estado; VM = $fila.VM; vCenter = $fila.vCenter
                Version = if ($infoVC) { $infoVC.Version } else { "Not Found" }
                Build = if ($infoVC) { $infoVC.Build } else { "Not Found" }
                "Sizing actual" = $fila."Sizing actual"
                "Sizing recomendado" = $fila."Sizing recomendado"
                "Cantidad de VMs" = $fila."Cantidad de VMs"
                "vCPU" = $fila."vCPU"; "Memory GB" = $fila."Memory GB"
            }
        }
    }

    $vcentersUnicos = $datosCombinados | Group-Object -Property @{ Expression = { $_.vCenter + '|' + $_.VM } } | ForEach-Object {
        $_.Group | Sort-Object -Property @{
            Expression = {
                $sizing = ""; if (-not [string]::IsNullOrEmpty($_."Sizing actual")) { $sizing = $_."Sizing actual".ToLower().Trim() }
                if ($sizingWeights.ContainsKey($sizing)) { return $sizingWeights[$sizing] } else { return 0 }
            }
        } | Select-Object -Last 1
    }

    $informeFinal = @()

    foreach ($vcenter in $vcentersUnicos) {
        $majorVersion = ($vcenter.Version -split '\.')[0]
        
        if ($sizingByVersion.ContainsKey($majorVersion)) {
            $requisitos = $sizingByVersion[$majorVersion].requisitos
        }

        $sizingActual = ""; if (-not [string]::IsNullOrEmpty($vcenter."Sizing actual")) { $sizingActual = $vcenter."Sizing actual".ToLower().Trim() }

        if ($requisitos.ContainsKey($sizingActual)) {
            $req = $requisitos[$sizingActual]
            
            $coresActuales = 0; $memoriaActual = 0.0; $vmsActuales = 0
            [int]::TryParse($vcenter.'vCPU', [ref]$coresActuales)
            [double]::TryParse(($vcenter.'Memory GB'.ToString()).Replace(',','.'), [ref]$memoriaActual)
            [int]::TryParse($vcenter.'Cantidad de VMs', [ref]$vmsActuales)

            if (($coresActuales -lt $req.Cores) -or ($memoriaActual -lt $req.MemoriaGB) -or ($vmsActuales -gt $req.VMsMax)) {
                $informeFinal += $vcenter
            }
        }
    }

    $resultadoChecklist = "No recomendado"
    $detalleChecklist = "Se encontraron $($informeFinal.Count) vCenters mal dimensionados"

    if ($informeFinal) {
        $datosParaExportar = $informeFinal | Select-Object -Property @(
            'State', 'VM', 'vCenter', 'Version', 'Build',
            'Sizing actual',
            @{ Name = 'VMs';          Expression = { $_.'Cantidad de VMs' } },
            @{ Name = 'Cores';        Expression = { $_.'vCPU' } },
            @{ Name = 'Memoria (GB)'; Expression = { $_.'Memory GB' } }
        )

        Exportar-InformeConEstilo -Datos $datosParaExportar `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Sizing Incorrecto" `
                                  -ColumnasSinConversion "VM", "vCenter", "Version"

    } else {
        $resultadoChecklist = "Resultado Esperado"
        $detalleChecklist = "Todos los vCenter Servers se encuentran dimensionados correctamente"
    }

    return [PSCustomObject]@{
        ID        = "VSP-ME-01"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}


function Particiones {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "Partitions"
        
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Disk'] -and $fila.PSObject.Properties['Free %']) {
                $valorNormalizado = ([string]$fila."Free %").Replace(',', '.')
                $valorLimpio = $valorNormalizado -replace '[^0-9.]'
                $freePercentNumeric = 0
                [double]::TryParse($valorLimpio, [ref]$freePercentNumeric)
                $diskPath = $fila.Disk.Trim()
                
                if (($freePercentNumeric -lt 30) -and ($diskPath -ne "/storage/core" -and $diskPath -ne "/storage/archive")) {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "VM"         = $fila."VM"
                        "Annotation" = $fila."Annotation"
                        "Disk"       = $fila."Disk"
                        "Free %"     = $fila."Free %"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Espacio en Particiones"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-02"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) particiones con espacio libre inferior al 30%"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-02"
            Resultado = "Resultado Esperado"
            Detalle   = "Todas las particiones de disco analizadas tienen un espacio libre superior al 30%"
        }
    }
}


function SyslogCheck {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['SyslogGlobalLogDir'] -and $fila.PSObject.Properties['SyslogGlobalLogHost']) {

                $logDir = $fila.SyslogGlobalLogDir
                $logHost = $fila.SyslogGlobalLogHost

                $dirEsIncorrecto = ($logDir -match "/scratch/log" -or $logDir -match "local" -or [string]::IsNullOrEmpty($logDir))
                $hostEsIncorrecto = ($logHost -match "/scratch/log" -or $logHost -match "local" -or [string]::IsNullOrEmpty($logHost))

                if ($dirEsIncorrecto -and $hostEsIncorrecto) {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"             = $fila."vCenter"
                        "Hostname"            = $fila."Hostname"
                        "Datacenter"          = $fila."Datacenter"
                        "Cluster"             = $fila."Cluster"
                        "SyslogGlobalLogDir"  = $logDir
                        "SyslogGlobalLogHost" = $logHost
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $datosSalida `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Syslogs y ubicacion de logs" `
                                  -ColumnasSinConversion "Hostname"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-04"
            Resultado = "No configurado"
            Detalle   = "Se encontraron $($datosSalida.Count) hosts sin configuracion de Syslog remoto"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-04"
            Resultado = "Resultado Esperado"
            Detalle   = "Todos los hosts tienen una configuracion de Syslog valida"
        }
    }
}


function Multipath {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "Datastores"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Datastore'] -and $fila.PSObject.Properties['Policy'] -and $fila.PSObject.Properties['Hostname']) {

                $datastoreName = $fila.Datastore
                $hostnameCorto = ($fila.Hostname -split '\.')[0]
                
                if ($datastoreName -match $hostnameCorto) {
                    continue 
                }

                $policy = $fila.Policy.Trim()

                if ($policy -eq 'vSAN') {
                    continue
                }

                $esLocal = ($datastoreName -match "local" -or $datastoreName -match "datastore1")
                $esCompartido = -not($esLocal)
                
                $policyRecomendada = ""
                $esInconsistente = $false

                if ($esCompartido -and $policy -ne "RoundRobin") {
                    $esInconsistente = $true
                    $policyRecomendada = "RoundRobin"
                }

                if ($esLocal -and $policy -eq "RoundRobin") {
                    $esInconsistente = $true
                    $policyRecomendada = "MRU o Fixed"
                }

                if ($esInconsistente) {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"             = $fila."vCenter"
                        "Hostname"            = $fila."Hostname"
                        "Datastore"           = $datastoreName
                        "Policy"              = $policy
                        "Policy recomendado"  = $policyRecomendada
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $datosSalida `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Politicas de Multipath" `
                                  -ColumnasSinConversion "Hostname"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-05"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) datastores con politicas de Multipath no recomendadas"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-05"
            Resultado = "Resultado Esperado"
            Detalle   = "Todos los datastores analizados cumplen con las politicas de Multipath recomendadas"
        }
    }
}


function ConsVer { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        
        $gruposCluster = @{}

        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                
                if (-not $gruposCluster.ContainsKey($claveUnica)) {
                    $gruposCluster[$claveUnica] = @()
                }
                $gruposCluster[$claveUnica] += $fila
            }
        }
        
        foreach ($clave in $gruposCluster.Keys) {
            $hostsDelCluster = $gruposCluster[$clave]
            
            $versionesUnicas = $hostsDelCluster | ForEach-Object { 
                if ($_.EsxiVersion) { $_.EsxiVersion.Trim() } 
            } | Select-Object -Unique
            
            if ($versionesUnicas.Count -gt 1) {
                $datosSalida += $hostsDelCluster
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        $informeFinal = $datosSalida | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"     = $_."vCenter"
                "Hostname"    = $_."Hostname"
                "Datacenter"  = $_."Datacenter"
                "Cluster"     = $_."Cluster"
                "EsxiVersion" = $_."EsxiVersion"
            }
        }
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Consistencias de versiones" `
                                  -ColumnasSinConversion "Hostname"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-06"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron clusteres con versiones de ESXi inconsistentes"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-06"
            Resultado = "Resultado Esperado"
            Detalle   = "Todos los clusteres analizados tienen versiones de ESXi consistentes"
        }
    }
}

function ConsRec {
    param(
        [double]$toleranciaMemoriaGB = 1,
        [double]$toleranciaVelocidadCPU = 100 
    ) 
    
    $datosSalida = @() 

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        $gruposCluster = @{}

        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                if (-not $gruposCluster.ContainsKey($claveUnica)) {
                    $gruposCluster[$claveUnica] = @()
                } 
                $gruposCluster[$claveUnica] += $fila 
            } 
        }

        foreach ($clave in $gruposCluster.Keys) { 
            $hostsDelCluster = $gruposCluster[$clave]
            
            $memorias = @(); $velocidades = @(); $modelos = @()

            foreach ($h in $hostsDelCluster) {
                $memStr = ([string]$h.MemoryGB) -replace '[^0-9\.,]', '' -replace ',', '.'
                $spdStr = ([string]$h.CpuSpeed) -replace '[^0-9\.,]', '' -replace ',', '.'
                $memorias += [double]$memStr 
                $velocidades += [double]$spdStr
                $modelos += $h.CpuModel.ToString().Trim()
            }

            $memMin = ($memorias | Measure-Object -Minimum).Minimum
            $memMax = ($memorias | Measure-Object -Maximum).Maximum
            $spdMin = ($velocidades | Measure-Object -Minimum).Minimum
            $spdMax = ($velocidades | Measure-Object -Maximum).Maximum
            $modelosUnicos = $modelos | Select-Object -Unique

            $inconsistente = $false
            if ($modelosUnicos.Count -gt 1) { $inconsistente = $true }
            if (($memMax - $memMin) -gt $toleranciaMemoriaGB) { $inconsistente = $true }
            if (($spdMax - $spdMin) -gt $toleranciaVelocidadCPU) { $inconsistente = $true }

            if ($inconsistente) {
                $datosSalida += $hostsDelCluster
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        $informeFinal = $datosSalida | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"  = $_."vCenter"
                "Hostname" = $_."Hostname"
                "Cluster"  = $_."Cluster"
                "MemoryGB" = $_."MemoryGB"
                "CpuModel" = $_."CpuModel"
                "CpuSpeed" = $_."CpuSpeed"
            }
        }
        Exportar-InformeConEstilo   -Datos $informeFinal `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Consistencias de recursos" `
                                    -ColumnasSinConversion "Hostname"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-07"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron clusteres con inconsistencias de recursos de hardware"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-07"
            Resultado = "Resultado Esperado"
            Detalle   = "Todos los clusteres analizados tienen recursos de hardware consistentes"
        }
    }
}


function PlacaRed {
    $vmsConE1000 = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "VM"

        foreach ($fila in $vPartition) {
            # Verificar si la fila y las propiedades necesarias existen
            if ($null -ne $fila `
                -and $fila.PSObject.Properties['SO (vCenter)'] `
                -and $fila.PSObject.Properties['VM']) {

                $soActual = $fila."SO (vCenter)".Trim()
                $vmName = $fila.VM # Obtenemos el nombre de la VM

                # Primer filtro: Sistema Operativo Windows (excluyendo versiones antiguas)
                if (($soActual -like "Microsoft Windows*") `
                    -and ($soActual -notlike "*2003*") `
                    -and ($soActual -notlike "*2000*")) {

                    # --- INICIO DE LA NUEVA CONDICIÓN DE FILTRADO ---
                    # Segundo filtro: El nombre de la VM NO debe contener "_replica", "_rep" o "_cont"
                    if ($vmName -notmatch "_replica|_rep|_cont") {
                    # --- FIN DE LA NUEVA CONDICIÓN DE FILTRADO ---

                        # Tercer filtro: Buscar adaptadores E1000/Flexible
                        $encontroE1000 = $false
                        for ($i = 1; $i -le 10; $i++) {
                            $nombreAdapter = "Adapter_{0:D2}" -f $i
                            if ($fila.PSObject.Properties[$nombreAdapter]) {
                                $valorAdapter = ([string]$fila.$nombreAdapter).Trim()
                                if ($valorAdapter -eq "e1000" -or $valorAdapter -eq "e1000e" -or $valorAdapter -eq "Flexible") {
                                    $encontroE1000 = $true
                                    break
                                }
                            }
                        }

                        # Si pasó todos los filtros (SO, nombre y adaptador), se añade
                        if ($encontroE1000) {
                            $vmsConE1000 += $fila
                        }
                    } # Cierre del if para el filtro de nombre
                } # Cierre del if para el filtro de SO
            } # Cierre del if para verificar propiedades
        } # Cierre del foreach ($fila in $vPartition)
    } # Cierre del Get-ChildItem | ForEach-Object

    # --- Lógica de la Checklist y Exportación (sin cambios) ---
    if ($vmsConE1000.Count -eq 0) {
        return [PSCustomObject]@{ ID = "VSP-ME-08"; Resultado = "Resultado Esperado"; Detalle = "No se encontraron VMs Windows (excluyendo réplicas) con adaptadores E1000/E1000e." }
    } else {
        $resultadoChecklist = ""
        $detalleChecklist = ""

        if ($vmsConE1000.ToolsStatus -contains 'toolsNotInstalled') {
            $resultadoChecklist = "VMware Tools no instaladas"; $detalleChecklist = "Se encontraron $($vmsConE1000.Count) VMs con adaptadores E1000, y al menos una no tiene VMware Tools instaladas."
        } elseif ($vmsConE1000 | Where-Object { $_.ToolsStatus -ne 'toolsOk' }) {
            $resultadoChecklist = "VMware Tools desactualizadas"; $detalleChecklist = "Se encontraron $($vmsConE1000.Count) VMs con adaptadores E1000, y al menos una no tiene VMware Tools actualizadas."
        } else {
            $resultadoChecklist = "VMware Tools OK"; $detalleChecklist = "Se encontraron $($vmsConE1000.Count) VMs con adaptadores E1000, pero todas tienen VMware Tools actualizadas."
        }

        $datosParaExportar = $vmsConE1000 | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"      = $_."vCenter"; "VM" = $_."VM"; "Cluster" = $_."Cluster"; "Host" = $_."Host"
                "State"        = $_."State"; "ToolsStatus" = $_."ToolsStatus"; "SO (vCenter)" = $_."SO (vCenter)"
                "SO (Tools)"   = $_."SO (Tools)"; "Adapter_01" = $_."Adapter_01"; "Adapter_02" = $_."Adapter_02"
                "Adapter_03"   = $_."Adapter_03"; "Adapter_04" = $_."Adapter_04"; "Adapter_05" = $_."Adapter_05"
                "Adapter_06"   = $_."Adapter_06"; "Adapter_07" = $_."Adapter_07"; "Adapter_08" = $_."Adapter_08"
                "Adapter_09"   = $_."Adapter_09"; "Adapter_10" = $_."Adapter_10"
            }
        }
        Exportar-InformeConEstilo -Datos $datosParaExportar `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Hardware Placa de Red" `
                                  -ColumnasSinConversion "Host"

        return [PSCustomObject]@{ ID = "VSP-ME-08"; Resultado = $resultadoChecklist; Detalle = $detalleChecklist }
    }
}

function TSMCheck { 
    $datosSalida = @()
    $encontroCero = $false
    $encontroFueraDeRango = $false

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['ESXIShellTimeOut'] -and $fila.PSObject.Properties['ESXIShellinteractiveTimeOut']) {

                $timeout = $fila.ESXIShellTimeOut
                $interactiveTimeout = $fila.ESXIShellinteractiveTimeOut
                $esIncorrecto = $false

                if ($timeout -eq 0 -or $interactiveTimeout -eq 0) {
                    $encontroCero = $true
                    $esIncorrecto = $true
                
                } elseif (($timeout -lt 300 -or $timeout -gt 1800) -or ($interactiveTimeout -lt 300 -or $interactiveTimeout -gt 1800)) {
                    $encontroFueraDeRango = $true
                    $esIncorrecto = $true
                }
                
                if ($esIncorrecto) {
                    $datosSalida += [PSCustomObject]@{
                        "vCenter"                     = $fila."vCenter"
                        "Hostname"                    = $fila."Hostname"
                        "Datacenter"                  = $fila."Datacenter"
                        "Cluster"                     = $fila."Cluster"
                        "ESXIShellTimeOut"            = $timeout
                        "ESXIShellinteractiveTimeOut" = $interactiveTimeout
                    }
                }
            }
        }
    }
    
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Todos los hosts tienen la configuracion de TSM dentro del rango recomendado (300-1800s)"
    
    if ($encontroCero) {
        $resultadoChecklist = "No recomendado"
        $detalleChecklist = "Se encontraron hosts con timeout de TSM configurado en 0 segundos"
    } elseif ($encontroFueraDeRango) {
        $resultadoChecklist = "Fuera de valor recomendado"
        $detalleChecklist = "Se encontraron hosts con timeout de TSM configurado fuera del rango recomendado (300-1800s)"
    }

    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Verificacion de TSM" `
                                    -ColumnasSinConversion "Hostname"
    }

    return [PSCustomObject]@{
        ID        = "VSP-ME-12"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}

function pManagement {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"

        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and $fila.PSObject.Properties['PowerManagement']) {
                if ($fila.PowerManagement.Trim() -ne "High performance") {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"         = $fila."vCenter"
                        "Hostname"        = $fila."Hostname"
                        "Datacenter"      = $fila."Datacenter"
                        "Cluster"         = $fila."Cluster"
                        "PowerManagement" = $fila."PowerManagement"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Power Management" `
                                    -ColumnasSinConversion "Hostname"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-14"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) hosts con una politica de Power Management no recomendada"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-14"
            Resultado = "Resultado Esperado"
            Detalle   = "Todos los hosts analizados tienen la politica de Power Management configurada en High Performance"
        }
    }
}


function vmtools {
    $vmsConProblemas = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vmsConProblemas += Import-Excel -Path $archivoEntrada -WorksheetName "VM"
    }

    $vmsConProblemas = $vmsConProblemas | Where-Object {
        # --- Filtro de Status (Existente) ---
        $status = $_.ToolsStatus.Trim()
        $state = $_.State.Trim()
        $statusProblematico = -not (($status -eq "toolsOk") -or (($state -eq "PoweredOff") -and ($status -eq "toolsNotRunning")))

        # --- Filtro de Nombre (Nuevo) ---
        # Excluir si el nombre de la VM contiene _replica, _rep, o _cont
        # Usamos -notmatch para verificar que el nombre NO contenga ninguna de esas cadenas
        $nombreValido = $_.VM -notmatch "(_replica|_rep|_cont)"
        
        # La VM debe cumplir ambas condiciones para ser incluida
        $statusProblematico -and $nombreValido
    }
    
    if ($vmsConProblemas.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-ME-15"
            Resultado = "Resultado Esperado"
            Detalle   = "Todas las VMs tienen las VMware Tools en estado 'OK'"
        }
    } else {
        $resultadoChecklist = ""
        $detalleChecklist = ""

        $clustersAfectados = $vmsConProblemas.Cluster | Select-Object -Unique

        $clusterEsInconsistente = $false
        foreach ($clusterNombre in $clustersAfectados) {
            $versionesToolsEnCluster = $vmsConProblemas | Where-Object { $_.Cluster -eq $clusterNombre } | Select-Object -ExpandProperty ToolsVersion -Unique
            
            if ($versionesToolsEnCluster.Count -gt 1) {
                $clusterEsInconsistente = $true
                break 
            }
        }

        if ($clusterEsInconsistente) {
            $resultadoChecklist = "Cluster no consistente"
            $detalleChecklist = "Se encontraron VMs con problemas de VMware Tools en clusteres con versiones de Tools inconsistentes"
        } else {
            $resultadoChecklist = "Cluster consistente"
            $detalleChecklist = "Se encontraron VMs con problemas de VMware Tools, y las versiones de Tools son consistentes dentro de cada cluster"
        }

        $datosParaExportar = $vmsConProblemas | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"              = $_."vCenter"
                "VM"                   = $_."VM"
                "Cluster"              = $_."Cluster"
                "Host"                 = $_."Host"
                "ConnectionState"      = $_."ConnectionState"
                "State"                = $_."State"
                "ToolsStatus"          = $_."ToolsStatus"
                "ToolsVersion"         = $_."ToolsVersion"
                "ToolsRequiredVersion" = $_."ToolsRequiredVersion"
            }
        }

        Exportar-InformeConEstilo   -Datos $datosParaExportar `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "VMware Tools" `
                                    -ColumnasSinConversion "Host"

        return [PSCustomObject]@{
            ID        = "VSP-ME-15"
            Resultado = $resultadoChecklist
            Detalle   = $detalleChecklist
        }
    }
}


function isos {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "VM"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['IsoConnected']) {
                $isoValue = $fila.IsoConnected
                
                if (-not [string]::IsNullOrEmpty($isoValue) -and $isoValue.Trim() -ne "[]") {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"      = $fila."vCenter"
                        "VM"           = $fila."VM"
                        "Cluster"      = $fila."Cluster"
                        "Host"         = $fila."Host"
                        "IsoConnected" = $fila."IsoConnected"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "ISOs Conectadas" `
                                    -ColumnasSinConversion "Host"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-16"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) VMs con ISOs conectadas"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-16"
            Resultado = "Resultado Esperado"
            Detalle   = "No se encontraron VMs con ISOs conectadas."
        }
    }
}

function placasDeRed {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "vNetwork"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Status'] -and $fila.PSObject.Properties['Connected'] -and $fila.PSObject.Properties['StartsConnected']) {
                if (([string]$fila.Status -eq "1") -and ($fila.Connected -eq "True") -and ($fila.StartsConnected -eq "False")) {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"         = $fila."vCenter"
                        "VM"              = $fila."VM"
                        "Cluster"         = $fila."Cluster"
                        "Host"            = $fila."Host"
                        "Status"          = $fila."Status"
                        "Mac"             = $fila."Mac"
                        "Connected"       = $fila."Connected"
                        "StartsConnected" = $fila."StartsConnected"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Placas No Inician Conectadas" `
                                    -ColumnasSinConversion "Host"
        
        return [PSCustomObject]@{
            ID        = "VSP-ME-17"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) placas de red no configuradas para iniciar conectadas"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-ME-17"
            Resultado = "Resultado Esperado"
            Detalle   = "Todas las placas de red de VMs estan configuradas para iniciar conectadas."
        }
    }
}


function snapshotsCheck {
    $datosSalida = @()

    $fechaLimite = (Get-Date).AddDays(-14)

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "Snapshot" | Where-Object {
            $_.VM -notmatch "(_replica|_rep|_cont)"
        } | Where-Object {
            $_.Snapshot -notmatch "Restore Point"
        }
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Fecha'] -and $fila.PSObject.Properties['SizeMB']) {
                try {
                    $fechaSnapshot = [datetime]$fila.Fecha
                    
                    if ($fechaSnapshot -lt $fechaLimite) {
                        $datosSalida += $fila
                    }
                } catch {
                    Write-Warning "No se pudo convertir la fecha '$($fila.Fecha)' para la VM '$($fila.VM)'. Se omite esta fila."
                }
            }
        }
    }

    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Snapshots Antiguos"

        $resultadoME18 = [PSCustomObject]@{
            ID        = "VSP-ME-18"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) snapshots con mas de 14 dias"
        }
    } else {
        $resultadoME18 = [PSCustomObject]@{
            ID        = "VSP-ME-18"
            Resultado = "Resultado Esperado"
            Detalle   = "No se encontraron snapshots con mas de 14 dias."
        }
    }

    $totalSizeMB = ($datosSalida | Measure-Object -Property SizeMB -Sum).Sum
    $totalSizeGB = [math]::Round($totalSizeMB / 1024, 2)

    $snapshotMasGrande = $datosSalida | Sort-Object -Property @{ Expression = { [double]$_.SizeMB } } | Select-Object -Last 1

    if ($snapshotMasGrande) {
        $sizeEnGB = [math]::Round(([double]$snapshotMasGrande.SizeMB / 1024), 2)

        if ([double]$snapshotMasGrande.SizeMB -gt (300 * 1024)) {
            $resultadoME03 = [PSCustomObject]@{
                ID        = "VSP-ME-03"
                Resultado = "No recomendado"
                Detalle   = "Se encontro un snapshot ('$($snapshotMasGrande.Snapshot)') de $($sizeEnGB)GB, que supera los 300GB. Total: $totalSizeGB GB"
            }
        } else {
            $resultadoME03 = [PSCustomObject]@{
                ID        = "VSP-ME-03"
                Resultado = "Resultado Esperado"
                Detalle   = "Ningun snapshot supera los 300GB. El mas grande es de $($sizeEnGB)GB. Total: $totalSizeGB GB"
            }
        }
    } else {
        $resultadoME03 = [PSCustomObject]@{
            ID        = "VSP-ME-03"
            Resultado = "Resultado Esperado"
            Detalle   = "No existen snapshots antiguos."
        }
    }

    return $resultadoME03, $resultadoME18
}

function endOfSupport {
    $datosSalida = @()
    $encontroVencido = $false
    $encontroPorVencer = $false

    $eosDates = @{
        '8.0' = '2027-10-11'
        '7.0' = '2025-10-02'
        '6.7' = '2022-10-15'
        '6.5' = '2022-10-15'
    }

    $fechaHoy = Get-Date
    $fechaLimiteUnAnio = $fechaHoy.AddYears(1)

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['ESXiVersion']) {
                $versionCompleta = $fila.ESXiVersion
                $versionPrincipal = $null

                if ($versionCompleta -match '(\d+\.\d+)') {
                    $versionPrincipal = $matches[1]
                }

                if ($versionPrincipal -and $eosDates.ContainsKey($versionPrincipal)) {
                    $fechaEos = [datetime]$eosDates[$versionPrincipal]
                    $esInvalido = $false

                    if ($fechaEos -lt $fechaHoy) {
                        $encontroVencido = $true
                        $esInvalido = $true
                    } elseif ($fechaEos -lt $fechaLimiteUnAnio) {
                        $encontroPorVencer = $true
                        $esInvalido = $true
                    }
                    
                    if ($esInvalido) {
                        $datosSalida += [PSCustomObject]@{
                            "vCenter"      = $fila."vCenter"
                            "Hostname"     = $fila."Hostname"
                            "Datacenter"   = $fila."Datacenter"
                            "Cluster"      = $fila."Cluster"
                            "Version"      = $versionCompleta
                            "Fecha de EoS" = $fechaEos.ToString('yyyy-MM-dd')
                        }
                    }
                }
            }
        }
    }
    
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Todos los productos tienen soporte vigente por más de un año."
    
    if ($encontroVencido) {
        $resultadoChecklist = "Vencido"
        $detalleChecklist = "Se encontraron productos fuera de soporte"
    } elseif ($encontroPorVencer) {
        $resultadoChecklist = "Vencimiento < 1 Año"
        $detalleChecklist = "Se encontraron productos con vencimiento de soporte en menos de 365 dias"
    }

    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "End of Support" `
                                    -ColumnasSinConversion "Hostname"
    }

    return [PSCustomObject]@{
        ID        = "VSP-TR-01"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}

function compatComponentes {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Supported']) {
                if ($fila.Supported.Trim() -eq "False") {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"                   = $fila."vCenter"
                        "Hostname"                  = $fila."Hostname"
                        "Datacenter"                = $fila."Datacenter"
                        "Cluster"                   = $fila."Cluster"
                        "Version"                   = $fila."ESXiVersion"
                        "Version minima soportada"  = $fila."Supported Releases"
                        "Supported"                 = $fila."Supported"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Compatibilidad de Componentes"
        
        return [PSCustomObject]@{
            ID        = "VSP-TR-02"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) componentes no soportados por la version de ESXi actual"
        }
    } else {
        return [PSCustomObject]@{
            ID        = "VSP-TR-02"
            Resultado = "Resultado Esperado"
            Detalle   = "Todos los componentes analizados son compatibles con su versión de ESXi."
        }
    }
}


function ntpCheck {
    $datosSalida = @()
    $encontroNoConfigurado = $false
    $encontroServicioDetenido = $false
    $encontroInconsistencia = $false

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        $gruposCluster = @{}

        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                if (-not $gruposCluster.ContainsKey($claveUnica)) {
                    $gruposCluster[$claveUnica] = @()
                }
                $gruposCluster[$claveUnica] += $fila
            }
        }
        
        foreach ($clave in $gruposCluster.Keys) {
            $hostsDelCluster = $gruposCluster[$clave]
            
            if ($hostsDelCluster.Where({ [string]::IsNullOrEmpty($_.NtpServer) })) { $encontroNoConfigurado = $true }
            if ($hostsDelCluster.Where({ $_.NtpdRunning -ne $true })) { $encontroServicioDetenido = $true }
            if (($hostsDelCluster.NtpServer | Select-Object -Unique).Count -gt 1) { $encontroInconsistencia = $true }

            $ntpsUnicos = $hostsDelCluster | ForEach-Object { if ($_) { $_.NtpServer.Trim() } } | Select-Object -Unique
            $hayInconsistencia = ($ntpsUnicos.Count -gt 1)
            $hayNtpdFalse = $hostsDelCluster | Where-Object { $_.NtpdRunning -ne $true }
            
            if ($hayInconsistencia) {
                $datosSalida += $hostsDelCluster
            }
            elseif ($hayNtpdFalse.Count -gt 0) {
                $datosSalida += $hayNtpdFalse
            }
        }
    }
    
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Todos los hosts tienen NTP configurado, en ejecucion y consistente por cluster."
    
    if ($encontroNoConfigurado) {
        $resultadoChecklist = "No configurado"
        $detalleChecklist = "Se encontraron hosts sin un servidor NTP configurado"
    } elseif ($encontroServicioDetenido) {
        $resultadoChecklist = "Configurado y servicio no activo"
        $detalleChecklist = "Se encontraron hosts con el servicio NTP detenido"
    } elseif ($encontroInconsistencia) {
        $resultadoChecklist = "Cluster no consistente"
        $detalleChecklist = "Se encontraron inconsistencias en la configuracion de NTP dentro de al menos un cluster"
    }

    if ($datosSalida.Count -gt 0) {
        $informeFinal = $datosSalida | Sort-Object -Property Hostname -Unique | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"     = $_."vCenter"
                "Hostname"    = $_."Hostname"
                "Datacenter"  = $_."Datacenter"
                "Cluster"     = $_."Cluster"
                "NtpServer"   = $_."NtpServer"
                "NtpdRunning" = $_."NtpdRunning"
            }
        }
        Exportar-InformeConEstilo   -Datos $informeFinal `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "NTP Problemas" `
                                    -ColumnasSinConversion "Hostname", "NtpServer"
    }

    return [PSCustomObject]@{
        ID        = "VSP-TR-03"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}


function Drivers { #en algun momento lo resolveremos    (de momento es unicamente informativo, no se filtra informacion de forma automatica o manual)
#    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
#        $archivoEntrada = $_.FullName
#    
#        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi IO"
#
#        $datosSalida += $vPartition | ForEach-Object {
#            $obj = [PSCustomObject]@{
#                "vCenter" = $_."vCenter"
#                "Hostname" = $_."Hostname"
#                "Version del Host" = $_."ESXi Release"
#                "Placa" = $_."Placa"
#                "Controlador" = $_."Controlador"
#                "Vendor" = $_."Vendor"
#                "Driver" = $_."Driver"
#                "Version" = $_."Version"
#                "Firmware" = $_."Firmware"
#                "Vid" = $_."Vid"
#                "Did" = $_."Did"
#                "Svid" = $_."Svid"
#                "ssid" = $_."ssid"
#                "URL" = $_."URL"
#            }
#            $obj
#        }
#    }
#    
#    if ($datosSalida) {
#        Exportar-InformeConEstilo   -Datos $datosSalida `
#                                    -RutaArchivo $archivoSalida `
#                                    -NombreHoja "Host Drivers" `
#                                    -ColumnasSinConversion "Hostname"}

    $resultadoChecklist = "Excpeción"
    $detalleChecklist = "No cuenta con filtrado de datos"

    return [PSCustomObject]@{
        ID        = "VSP-TR-04"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}


function DNSConfig { #modificar dropdown de checklist

    $datosSalida = @()
    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        $gruposCluster = @{}

        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                if (-not $gruposCluster.ContainsKey($claveUnica)) {
                    $gruposCluster[$claveUnica] = @()
                }
                $gruposCluster[$claveUnica] += $fila
            }
        }
        foreach ($clave in $gruposCluster.Keys) {
            $hostsDelCluster = $gruposCluster[$clave]
            
            $hostsSinDns = $hostsDelCluster.Where({ [string]::IsNullOrEmpty($_.DnsServer) })
            $hayDnsVacio = $hostsSinDns.Count -gt 0

            $dnsConfigurados = $hostsDelCluster.DnsServer.Trim() | Select-Object -Unique
            $hayInconsistencia = $dnsConfigurados.Count -gt 1
            
            if ($hayDnsVacio -or $hayInconsistencia) {
                $datosSalida += $hostsDelCluster
            }
        }
    }
    if ($datosSalida) {
        $informeFinal = $datosSalida | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"         = $_."vCenter"
                "Hostname"        = $_."Hostname"
                "Datacenter"      = $_."Datacenter"
                "Cluster"         = $_."Cluster"
                "ConnectionState" = $_."ConnectionState"
                "DnsServer"       = $_."DnsServer"
            }
        }
        Exportar-InformeConEstilo   -Datos $informeFinal `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Host DNS" `
                                    -ColumnasSinConversion "Hostname"

        $resultadoChecklist = "Cluster no consistente"
        $detalleChecklist = "Al menos un host no tiene un servidor DNS configurado, o no hay consistencia en un cluster"

        return [PSCustomObject]@{
            ID        = "VSP-TR-05"
            Resultado = $resultadoChecklist
            Detalle   = $detalleChecklist
        }
    }else {
        $resultadoChecklist = "Resultado Esperado"
        $detalleChecklist = "Todos los host tienen servidores DNS configurados y son consistenetes por cluster"

        return [PSCustomObject]@{
            ID        = "VSP-TR-05"
            Resultado = $resultadoChecklist
            Detalle   = $detalleChecklist
        }
    }
}

function Licencia {
    $datosSalida = @()
    $encontroVencida = $false
    $encontroPorVencer = $false
    $encontroEvaluation = $false
    $encontroExcedida = $false
    $encontroDuplicada = $false

    $totalLicencias = 0
    $licenciasEvaluationCount = 0
    $fechaHoy = Get-Date
    $fechaLimiteUnAnio = $fechaHoy.AddYears(1)

    $todasLasLicencias = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "vLicense"
    }

    # 1. Analizamos cada licencia para vencimiento, uso y evaluación
    foreach ($licencia in $todasLasLicencias) {
        $totalLicencias++
        $expiraRaw = $licencia.ExpirationDate
        $esProblematicaInicial = $false

        if ($licencia.Used -gt $licencia.Total) {
            $encontroExcedida = $true
            $esProblematicaInicial = $true
        }
        if ($expiraRaw -eq "Evaluation" -or $expiraRaw -eq "Never") {
            $encontroEvaluation = $true
            $licenciasEvaluationCount++
            $esProblematicaInicial = $true
        } else {
            try {
                $expiraConvertida = [datetimeoffset]::Parse($expiraRaw).DateTime
                if ($expiraConvertida -lt $fechaHoy) {
                    $encontroVencida = $true
                    $esProblematicaInicial = $true
                } elseif ($expiraConvertida -lt $fechaLimiteUnAnio) {
                    $encontroPorVencer = $true
                    $esProblematicaInicial = $true
                }
            } catch { # Silencioso
            }
        }
        if ($esProblematicaInicial) {
            $datosSalida += $licencia
        }
    }

    $gruposPorKey = $todasLasLicencias | Group-Object -Property LicenseKey
    $licenciasDuplicadas = $gruposPorKey | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Group }
    if ($licenciasDuplicadas.Count -gt 0) {
        $encontroDuplicada = $true
        $datosSalida += $licenciasDuplicadas
    }

    $resultadoChecklist = "Resultado Esperado"
    # ... (resto de la lógica if/elseif para $resultadoChecklist y $detalleChecklist) ...
    if ($encontroVencida) {
        $resultadoChecklist = "Vencido"
        $detalleChecklist = "Se encontraron licencias vencidas. Ver anexo."
    } elseif ($encontroPorVencer) {
        $resultadoChecklist = "Vencimiento < 1 Año"
        $detalleChecklist = "Se encontraron licencias con vencimiento menor a un año. Ver anexo."
    } elseif ($encontroExcedida) {
        $resultadoChecklist = "No recomendado"
        $detalleChecklist = "Se encontraron licencias cuyo uso excede el límite total. Ver anexo."
    } elseif ($encontroDuplicada) {
        $resultadoChecklist = "No recomendado"
        $detalleChecklist = "Se encontraron License Keys duplicadas. Ver anexo."
    } elseif ($encontroEvaluation -and $licenciasEvaluationCount -eq $totalLicencias) {
        $resultadoChecklist = "En evaluación"
        $detalleChecklist = "Todas las licencias del entorno se encuentran en modo de evaluación. Ver anexo."
    } elseif ($encontroEvaluation) {
        $resultadoChecklist = "No recomendado"
        $detalleChecklist = "Se encontraron licencias en modo de evaluación junto con otras licencias. Ver anexo."
    }


    # 4. Generación del Anexo (si hay problemas)
    if ($datosSalida.Count -gt 0) {
        $datosSalidaUnicos = $datosSalida | Group-Object -Property vCenter, LicenseKey, Name | ForEach-Object {
            $_.Group | Select-Object -First 1
        }

        $informeFinal = $datosSalidaUnicos | Sort-Object -Property Name, LicenseKey | ForEach-Object {
            [PSCustomObject]@{
                Licencia           = $_.Name
                LicenseKey         = $_.LicenseKey
                Componente         = $_.ProductName
                vCenter            = $_.vCenter
                "En Uso"           = $_.Used
                Total              = $_.Total
                "Fecha Expiracion" = $_.ExpirationDate
            }
        }
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Licencias Problemas"
    }

    # 5. Retorno de estado para la Checklist
    return [PSCustomObject]@{
        ID        = "VSP-TR-06"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}


function NIOC {
    $todosLosVDS = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "vDS"
    }

    $datosSalida = $todosLosVDS | Where-Object { $_."NIOC Enabled" -ne $true }
    
    if ($datosSalida.Count -gt 0) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "NIOC Deshabilitado"
        
        return [PSCustomObject]@{
            ID        = "VSP-TR-11"
            Resultado = "No recomendado"
            Detalle   = "Se encontraron $($datosSalida.Count) vDS con NIOC deshabilitado"
        }
    } else {
        $detalleChecklist = ""
        if ($todosLosVDS.Count -gt 0) {
            $detalleChecklist = "Todos los vDS encontrados tienen NIOC habilitado"
        } else {
            $detalleChecklist = "No se encontraron vDS en el entorno"
        }
        
        return [PSCustomObject]@{
            ID        = "VSP-TR-11"
            Resultado = "Resultado Esperado"
            Detalle   = $detalleChecklist
        }
    }
}


function vss {

    $informeFinal = @()

    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "Standard Switch"
    } | Where-Object { -not [string]::IsNullOrEmpty($_.Cluster) }

    $gruposCluster = $todosLosDatos | Group-Object -Property @{ Expression = { $_.vCenter + '|' + $_.Cluster } }

    foreach ($cluster in $gruposCluster) {
        
        $hostsUnicos = $cluster.Group.ESXi | Select-Object -Unique
        if ($hostsUnicos.Count -le 1) {
            continue
        }
        
        $gruposPorHost = $cluster.Group | Group-Object -Property ESXi
        
        $huellasDePortGroups = $gruposPorHost | ForEach-Object {
            ($_.Group.PortGroup | Select-Object -Unique | Sort-Object) -join ';'
        } | Select-Object -Unique
        
        $listaDePortGroupsEsInconsistente = $huellasDePortGroups.Count -gt 1

        $configuracionEsInconsistente = $false
        if (-not $listaDePortGroupsEsInconsistente) {
            $portgroupsEnCluster = $cluster.Group | Group-Object -Property PortGroup
            foreach ($pg in $portgroupsEnCluster) {
                $configuraciones = $pg.Group | ForEach-Object { "$($_.Switch)|$($_.vLAN)" } | Select-Object -Unique
                if ($configuraciones.Count -gt 1) {
                    $configuracionEsInconsistente = $true
                    break
                }
            }
        }
        
        if ($listaDePortGroupsEsInconsistente -or $configuracionEsInconsistente) {
            $informeFinal += $cluster.Group
        }
    }

    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "La configuracion de PortGroups es consistente entre los host del mismo cluster"

    if ($informeFinal) {
        $informeLimpio = $informeFinal | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"   = $_."vCenter"
                "ESXi"      = $_."ESXi"
                "Cluster"   = $_."Cluster"
                "PortGroup" = $_."PortGroup"
                "Switch"    = $_."Switch"
                "vLAN"      = $_."vLAN"
            }
        }
        
        $resultadoChecklist = "No recomendado"
        $detalleChecklist = "Configuracion de PortGroups no es consistente entre hosts del mismo cluster"

        Exportar-InformeConEstilo   -Datos $informeLimpio `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "VSS PortGroups"
    }

    return [PSCustomObject]@{
            ID        = "VSP-TR-12"
            Resultado = $resultadoChecklist
            Detalle   = $detalleChecklist
    }


}

function vCenterRoot {
    # 1. Importar datos de la hoja "vCenter"
    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "vCenter"
    } | Where-Object { -not [string]::IsNullOrEmpty($_. "vCenter Server") }

    # --- CONDICIÓN GLOBAL: SIN ACCESO (Hoja Vacía) ---
    if (-not $todosLosDatos -or $todosLosDatos.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-TR-09"
            Resultado = "Sin acceso"
            Detalle   = "No se pudo recolectar información del VAMI o la hoja 'vCenter' está vacía."
        }
    }

    $informeFinal = @()
    $hoy = Get-Date
    $fechaLimite6Meses = $hoy.AddMonths(6)
    
    # Banderas para el estado global (Solo lógica de fechas)
    $hayVencidos = $false
    $hayPorVencer = $false # < 6 meses

    # 2. Procesar cada fila
    foreach ($fila in $todosLosDatos) {
        $fechaStr = $fila."Expiration Date"
        $statusFinal = "Desconocido"
        $esHallazgo = $false 

        # --- Lógica de Negocio ---
        if ($fechaStr -eq "Nunca") {
            $statusFinal = "Resultado Esperado"
        }
        elseif ($fechaStr -eq "N/A" -or $fila.Status -like "*Error*" -or $fila.Status -like "*Unknown*") {
            # Error puntual en esta fila, pero no es "Sin acceso" global
            $statusFinal = "Error de lectura"
            $esHallazgo = $true
        }
        else {
            try {
                $fechaVencimiento = Get-Date $fechaStr -ErrorAction Stop
                
                if ($fechaVencimiento -lt $hoy) {
                    $statusFinal = "Vencido"
                    $esHallazgo = $true
                    $hayVencidos = $true
                }
                elseif ($fechaVencimiento -lt $fechaLimite6Meses) {
                    $statusFinal = "Vencimiento < 6 meses"
                    $esHallazgo = $true
                    $hayPorVencer = $true
                }
                else {
                    $statusFinal = "Resultado Esperado"
                }
            }
            catch {
                $statusFinal = "Error formato fecha"
                $esHallazgo = $true
            }
        }

        # 3. Si es un hallazgo, lo agregamos al informe detallado (Anexo)
        if ($esHallazgo) {
            $informeFinal += [PSCustomObject]@{
                "vCenter"         = $fila."vCenter Server"
                "Root User"       = $fila."Root User"
                "Expiration Date" = $fechaStr
                "Days Remaining"  = $fila."Days Remaining"
                "Status Original" = $fila."Status"
                "Evaluación"      = $statusFinal
            }
        }
    }

    # 4. Definir el Resultado Final del Checklist (Prioridad estricta)
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Las contraseñas de root no expiran o vencen en más de 6 meses."

    if ($hayVencidos) {
        $resultadoChecklist = "Vencido"
        $detalleChecklist = "Se detectaron contraseñas de root ya expiradas."
    }
    elseif ($hayPorVencer) {
        $resultadoChecklist = "Vencimiento < 6 meses"
        $detalleChecklist = "Se detectaron contraseñas próximas a vencer (menos de 6 meses)."
    }
    # Nota: Si solo hubo "Error de lectura" en filas puntuales, el estado se mantiene en "Resultado Esperado" 
    # (o podrías definir un estado intermedio si quisieras, pero según tu pedido, solo hay 3 estados válidos si hay datos).

    # 5. Exportar al Anexo Técnico (Si corresponde)
    if ($informeFinal.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Vencimiento password de Root"
    }

    # 6. Retornar objeto para el Checklist
    return [PSCustomObject]@{
        ID        = "VSP-TR-09"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}

# la siguiente funcion debera tomar de la hoja de Certficate la columna de vCenter, Ubicacion, Subject, Valid from, Valid Until y Emisor
# debera filtrar:
#Resultados:
#Si el vencimiento es igual o mayor a un año, el chequeo debe cerrarse con Resultado Esperado.
#Si el vencimiento es igual o mayor a 6 meses, deberá marcarse como No Esperado y notificarse en el informe proactivo.
#Si el vencimiento es menor a 6 meses, deberá generarse un ticket de soporte, tipo problema, prioridad planificado para dar seguimiento al caso.
function vCenterCert {
    # 1. Importar datos de la hoja "Certficate" (Nombre exacto solicitado)
    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "Certficate"
    } | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) }

    # Variables de control
    $informeFinal = @()
    $hoy = Get-Date
    $fechaUnAnio = $hoy.AddYears(1)
    
    # Banderas para determinar el estado global
    $hayVencidos = $false
    $hayMenorUnAnio = $false
    $hayErrores = $false

    # 2. Verificar si hay datos (Caso "Sin acceso")
    if (-not $todosLosDatos -or $todosLosDatos.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-SEC-02"
            Resultado = "Sin acceso"
            Detalle   = "No se pudo recolectar información de certificados o la hoja 'Certficate' está vacía."
        }
    }

    # 3. Procesar cada fila
    foreach ($fila in $todosLosDatos) {
        $fechaStr = $fila."Valid Until"
        $statusAnalisis = "Desconocido"
        $esHallazgo = $false

        if ([string]::IsNullOrEmpty($fechaStr) -or $fechaStr -eq "N/A") {
            $statusAnalisis = "Error al obtener fecha"
            $esHallazgo = $true
            $hayErrores = $true
        }
        else {
            try {
                $fechaVencimiento = Get-Date $fechaStr -ErrorAction Stop
                $diasRestantes = ($fechaVencimiento - $hoy).Days

                if ($fechaVencimiento -lt $hoy) {
                    # Ya expiró
                    $statusAnalisis = "Vencido"
                    $esHallazgo = $true
                    $hayVencidos = $true
                }
                elseif ($fechaVencimiento -lt $fechaUnAnio) {
                    # Vence en menos de 1 año (incluye los < 6 meses)
                    $statusAnalisis = "Vencimiento < 1 Año"
                    $esHallazgo = $true
                    $hayMenorUnAnio = $true
                }
                else {
                    # Vence en más de 1 año
                    $statusAnalisis = "Resultado Esperado"
                }
            }
            catch {
                $statusAnalisis = "Error formato fecha"
                $esHallazgo = $true
                $hayErrores = $true
            }
        }

        # Guardamos en el Anexo solo si no es el resultado ideal
        if ($esHallazgo) {
            $informeFinal += [PSCustomObject]@{
                "vCenter"       = $fila.vCenter
                "Ubicación"     = $fila.Ubicacion
                "Subject"       = $fila.Subject
                "Emisor"        = $fila.Emisor
                "Valid From"    = $fila."Valid From"
                "Valid Until"   = $fechaStr
                "Días Restantes"= $diasRestantes
            }
        }
    }

    # 4. Definir el Resultado Final del Checklist (Prioridad de gravedad)
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Todos los certificados tienen vigencia mayor a 1 año."

    if ($hayVencidos) {
        $resultadoChecklist = "Vencido"
        $detalleChecklist = "Se detectaron certificados expirados."
    }
    elseif ($hayErrores) {
        # Si no hay vencidos pero hubo errores de lectura, es una alerta
        $resultadoChecklist = "Sin acceso" 
        $detalleChecklist = "Hubo errores al leer las fechas de algunos certificados."
    }
    elseif ($hayMenorUnAnio) {
        $resultadoChecklist = "Vencimiento < 1 Año"
        $detalleChecklist = "Se detectaron certificados próximos a vencer (menos de 1 año)."
    }

    # 5. Exportar al Anexo Técnico (Si corresponde)
    if ($informeFinal.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Certificados vCenter"
    }

    # 6. Retorno para el Checklist
    return [PSCustomObject]@{
        ID        = "VSP-TR-07"
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}

function esxiCert {
    # 1. Importar datos de la hoja "ESXi"
    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "ESXi"
    } | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) }

    if (-not $todosLosDatos -or $todosLosDatos.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-SEC-03" 
            Resultado = "Sin acceso"
            Detalle   = "No se pudo recolectar información de los Hosts ESXi o la hoja está vacía."
        }
    }

    $informeFinal = @()
    $hoy = Get-Date
    $fechaUnAnio = $hoy.AddYears(1)
    
    # Contadores y acumuladores para el resumen
    $countVencidos = 0
    $countMenorUnAnio = 0
    $fechaMasProxima = $null # Para guardar la fecha más cercana a vencer

    # 2. Procesar cada fila
    foreach ($fila in $todosLosDatos) {
        $fechaStr = $fila."Cert Valid To"
        $statusAnalisis = "Desconocido"
        $esHallazgo = $false

        if ([string]::IsNullOrEmpty($fechaStr) -or $fechaStr -eq "N/A") {
            $statusAnalisis = "Error de lectura"
            $esHallazgo = $true
        }
        else {
            try {
                $fechaVencimiento = Get-Date $fechaStr -ErrorAction Stop
                
                if ($fechaVencimiento -lt $hoy) {
                    $statusAnalisis = "Vencido"
                    $esHallazgo = $true
                    $countVencidos++
                }
                elseif ($fechaVencimiento -lt $fechaUnAnio) {
                    $statusAnalisis = "Vencimiento < 1 Año"
                    $esHallazgo = $true
                    $countMenorUnAnio++
                    
                    # Logica para encontrar la fecha más próxima
                    if ($null -eq $fechaMasProxima -or $fechaVencimiento -lt $fechaMasProxima) {
                        $fechaMasProxima = $fechaVencimiento
                    }
                }
                else {
                    $statusAnalisis = "Resultado Esperado"
                }
            }
            catch {
                $statusAnalisis = "Error formato fecha"
                $esHallazgo = $true
            }
        }

        if ($esHallazgo) {
            $informeFinal += [PSCustomObject]@{
                "vCenter"       = $fila.vCenter
                "Hostname"      = $fila.Hostname
                "Cert Valid To" = $fechaStr
                "Cert Issuer"   = $fila."Cert Issuer"
            }
        }
    }

    # 4. Definir el Resultado Final y Detalle Rico
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Todos los certificados de host tienen vigencia mayor a 1 año."

    if ($countVencidos -gt 0) {
        $resultadoChecklist = "Vencido"
        # Detalle específico: Cantidad de vencidos
        $detalleChecklist = "Se detectaron $countVencidos certificados de host ESXi ya expirados."
    }
    elseif ($countMenorUnAnio -gt 0) {
        $resultadoChecklist = "Vencimiento < 1 Año"
        
        # Formateamos la fecha más próxima para que se vea bien
        $fechaProximaStr = if ($fechaMasProxima) { $fechaMasProxima.ToString("yyyy-MM-dd") } else { "N/A" }
        
        # Detalle específico: Cantidad y fecha más próxima
        $detalleChecklist = "Se detectaron $countMenorUnAnio certificados próximos a vencer. El más próximo vence el $fechaProximaStr."
    }

    # 5. Exportar al Anexo Técnico
    if ($informeFinal.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Certificados ESXi"
    }

    # 6. Retornar objeto para el Checklist
    return [PSCustomObject]@{
        ID        = "VSP-TR-08" 
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}


function performanceHealthCheck {
    # 1. Importar datos de la hoja "PerformanceHealth"
    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "PerformanceHealth"
    } | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) }

    # --- CONDICIÓN GLOBAL: SIN DATOS ---
    if (-not $todosLosDatos -or $todosLosDatos.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-MON-01" # Ajustar ID
            Resultado = "Sin acceso"
            Detalle   = "No se pudo recolectar información de salud de performance o la hoja está vacía."
        }
    }

    $informeFinal = @()
    $hayProblemas = $false
    $clustersAfectados = 0

    # 2. Procesar cada fila (cada Cluster)
    foreach ($fila in $todosLosDatos) {
        $esHallazgo = $false
        $motivoFallo = @()

        # Verificamos cada columna crítica
        if ($fila."Health DB" -ne "OK") {
            $esHallazgo = $true
            $motivoFallo += "Fallo DB General"
        }
        if ($fila."CPU Stats" -ne "OK") {
            $esHallazgo = $true
            $motivoFallo += "Falta CPU"
        }
        if ($fila."Mem Stats" -ne "OK") {
            $esHallazgo = $true
            $motivoFallo += "Falta Memoria"
        }
        if ($fila."Net Stats" -ne "OK") {
            $esHallazgo = $true
            $motivoFallo += "Falta Red"
        }
        if ($fila."Disk Stats" -ne "OK") {
            $esHallazgo = $true
            $motivoFallo += "Falta Disco"
        }

        # 3. Si hay problemas, guardamos para el reporte detallado
        if ($esHallazgo) {
            $hayProblemas = $true
            $clustersAfectados++
            
            $informeFinal += [PSCustomObject]@{
                "vCenter"     = $fila.vCenter
                "Cluster"     = $fila.Cluster
                "Estado DB"   = $fila."Health DB"
                "CPU"         = $fila."CPU Stats"
                "Memoria"     = $fila."Mem Stats"
                "Red"         = $fila."Net Stats"
                "Disco"       = $fila."Disk Stats"
            }
        }
    }

    # 4. Definir el Resultado Final del Checklist
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Todos los clústeres están recolectando y almacenando métricas de rendimiento correctamente."

    if ($hayProblemas) {
        $resultadoChecklist = "No recomendado"
        $detalleChecklist = "Se detectaron $clustersAfectados clústeres con problemas en la recolección de estadísticas."
    }

    # 5. Exportar al Anexo Técnico (Si corresponde)
    if ($informeFinal.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Salud Performance"
    }

    # 6. Retornar objeto para el Checklist
    return [PSCustomObject]@{
        ID        = "VSP-TR-10" 
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}

function alarmCheck {
    # 1. Importar datos
    # Usamos la misma lógica que en performanceHealthCheck que confirmaste que funciona
    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        try { 
            Import-Excel -Path $_.FullName -WorksheetName "Falso Positivo" -ErrorAction Stop 
        } catch { 
            # Si la hoja no existe en un excel viejo, lo ignoramos
        }
    } | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) }

    # --- CONDICIÓN GLOBAL: SIN ACCESO (Si no se trajo nada de ningún lado) ---
    if (-not $todosLosDatos -or $todosLosDatos.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-ME-13" 
            Resultado = "Sin acceso"
            Detalle   = "No se encontraron registros de la prueba de alarmas (hoja vacía o no encontrada)."
        }
    }

    $informeFinal = @()
    $countExitos = 0
    $countFallas = 0
    $causas = @()
    $tokensEncontrados = @()

    # 2. Procesar cada fila
    foreach ($fila in $todosLosDatos) {
        # Limpieza y Normalización
        $resultado = if ($fila.Result) { $fila.Result.ToString().Trim() } else { "" }
        
        $nombreAlarma = if ($fila."Alarm Name") { $fila."Alarm Name" } else { $fila."Alarm Source" }
        
        $rutaScript = if ($fila."Script Path") { $fila."Script Path" } else { $fila."Alarm Path" }
        if (-not $rutaScript) { $rutaScript = $fila."Path Alarma" }

        $esHallazgo = $false
        $diagnostico = ""

        # --- VALIDACIÓN DE ÉXITO ---
        # Aceptamos OK, SUCCESS, o mensajes compuestos que empiezan con SUCCESS
        if ($resultado -eq "OK" -or $resultado -match "^SUCCESS") {
            $countExitos++

            # Análisis de Tokens (Solo en los exitosos)
            if (-not [string]::IsNullOrEmpty($rutaScript) -and $rutaScript -ne "N/A") {
                $partes = $rutaScript -split "\s+"
                if ($partes.Count -gt 1) {
                    $token = ($partes[1..($partes.Count - 1)]) -join " "
                    $tokensEncontrados += $token
                }
            }
        }
        else {
            # CASO FALLO
            $countFallas++
            $esHallazgo = $true
            
            # Diagnóstico Inteligente
            switch -Wildcard ($resultado) {
                "*No Script Configured*" { 
                    $diagnostico = "Configuración: Alarma sin acción de script"
                    if ("Configuración incompleta" -notin $causas) { $causas += "Configuración incompleta" }
                }
                "*Script Action Empty*" { 
                    $diagnostico = "Configuración: Acción de script vacía"
                    if ("Configuración incompleta" -notin $causas) { $causas += "Configuración incompleta" }
                }
                "*Not Found*" { 
                    $diagnostico = "Alarma no encontrada en vCenter"
                    if ("Alarmas faltantes" -notin $causas) { $causas += "Alarmas faltantes" }
                }
                "*No hay hosts*" { 
                    $diagnostico = "Sin hosts conectados"
                    if ("Entorno no disponible" -notin $causas) { $causas += "Entorno no disponible" }
                }
                "*Failed*" { 
                    $diagnostico = "Fallo funcional (Smoke Test)"
                    if ("Fallo funcional" -notin $causas) { $causas += "Fallo funcional" }
                }
                Default { 
                    $diagnostico = "Error: $resultado"
                    if ("Errores varios" -notin $causas) { $causas += "Errores varios" }
                }
            }

            # Guardar en Anexo
            $informeFinal += [PSCustomObject]@{
                "vCenter"       = $fila.vCenter
                "Host"          = $fila.Host
                "Alarma"        = $nombreAlarma
                "Script Path"   = $rutaScript
                "Resultado"     = $resultado
                "Diagnóstico"   = $diagnostico
                "Timestamp"     = $fila.Timestamp
            }
        }
    }

    # 3. Definir Resultado del Checklist
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Se verificaron $countExitos alarmas/pruebas correctamente."

    # Análisis de Consistencia de Tokens
    $tokensUnicos = @($tokensEncontrados | Select-Object -Unique)
    $notaInconsistencia = ""
    
    if ($tokensUnicos.Count -gt 1) {
        $notaInconsistencia = " (Nota: Se detectaron $($tokensUnicos.Count) variantes de tokens/argumentos)."
        # Si quieres agregar la nota al detalle de éxito:
        if ($countFallas -eq 0) { $detalleChecklist += $notaInconsistencia }
    }

    if ($countFallas -gt 0) {
        $resultadoChecklist = "No recomendado"
        $resumenCausas = $causas -join ", "
        $detalleChecklist = "Se detectaron $countFallas anomalías. Causas: $resumenCausas.$notaInconsistencia"
    }

    # 4. Exportar al Anexo Técnico
    if ($informeFinal.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Falso Positivo"
    }

    # 5. Retornar objeto
    return [PSCustomObject]@{
        ID        = "VSP-ME-13" 
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}


function backupCheck {
    # 1. Importar datos
    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "BackupActivity"
    } | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) }

    # Condición Global: Sin Datos
    if (-not $todosLosDatos -or $todosLosDatos.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-SE-02" 
            Resultado = "Sin acceso"
            Detalle   = "No se encontraron registros de actividad de backup."
        }
    }

    $informeFinal = @()
    
    # Banderas Globales para Checklist
    $hayFallidos = $false
    $hayManuales = $false
    $hayHuecos = $false
    $vcentersConProblemas = @()

    # 2. Agrupar por vCenter
    $gruposVcenter = $todosLosDatos | Group-Object vCenter

    foreach ($grupo in $gruposVcenter) {
        $nombreVcenter = $grupo.Name
        # Ordenamos descendente: índice 0 es el más reciente
        $backups = $grupo.Group | Sort-Object StartTime -Descending 
        
        $vcenterTieneProblemas = $false
        $mensajeProblema = ""

        # --- A. ANÁLISIS DE INTEGRIDAD Y ESTADO ---
        # Primero revisamos si hay fallos o manuales para levantar la bandera
        foreach ($b in $backups) {
            if ($b.Status -ne "SUCCEEDED") { $vcenterTieneProblemas = $true; $hayFallidos = $true }
            if ($b.Type -ne "SCHEDULED")   { $vcenterTieneProblemas = $true; $hayManuales = $true }
        }

        # --- B. ANÁLISIS DE FRECUENCIA (SECUENCIAL) ---
        $fechasParseadas = @()
        
        # 1. Parsear todas las fechas primero
        foreach ($b in $backups) {
            $fechaObj = $null
            try { 
                if ($b.StartTime -and $b.StartTime -ne "N/A") { 
                    $fechaObj = Get-Date $b.StartTime -ErrorAction Stop
                }
            } catch {}
            
            # Guardamos un objeto custom con la data original y la fecha parseada
            $fechasParseadas += [PSCustomObject]@{
                Original = $b
                DateObj  = $fechaObj
                Note     = "" # Para anotar saltos
            }
        }

        # 2. Verificar Hueco Inicial (Hoy vs Último Backup)
        $hoy = Get-Date
        if ($fechasParseadas.Count -gt 0 -and $fechasParseadas[0].DateObj) {
            $diffHoy = ($hoy - $fechasParseadas[0].DateObj).TotalHours
            if ($diffHoy -gt 26) {
                $vcenterTieneProblemas = $true
                $hayHuecos = $true
                $fechasParseadas[0].Note = "Salto > 24h desde hoy ($([math]::Round($diffHoy))hs)"
            }
        } else {
             # Si no hay fechas válidas en absoluto
             $vcenterTieneProblemas = $true
        }

        # 3. Verificar Huecos entre Backups (N vs N+1)
        for ($i = 0; $i -lt ($fechasParseadas.Count - 1); $i++) {
            $actual = $fechasParseadas[$i]
            $siguiente = $fechasParseadas[$i+1]

            if ($actual.DateObj -and $siguiente.DateObj) {
                $diffHoras = ($actual.DateObj - $siguiente.DateObj).TotalHours
                
                # Si la diferencia es mayor a 26 horas, hay un día faltante entre medio
                if ($diffHoras -gt 26) {
                    $vcenterTieneProblemas = $true
                    $hayHuecos = $true
                    # Marcamos el registro más antiguo del par para indicar que después de ese hubo un hueco
                    $siguiente.Note = "Salto > 24h hasta el siguiente backup ($([math]::Round($diffHoras))hs)"
                }
            }
        }

        # --- C. REPORTE (Si hay CUALQUIER problema, volcamos TODO el historial) ---
        if ($vcenterTieneProblemas) {
            if ($nombreVcenter -notin $vcentersConProblemas) { $vcentersConProblemas += $nombreVcenter }

            foreach ($item in $fechasParseadas) {
                $fila = $item.Original
                $notaFrecuencia = $item.Note
                
                # Determinamos la evaluación de esta fila específica
                $eval = "OK"
                if ($fila.Status -ne "SUCCEEDED") { $eval = "Estado Incorrecto ($($fila.Status))" }
                elseif ($fila.Type -ne "SCHEDULED") { $eval = "Tipo Incorrecto ($($fila.Type))" }
                elseif ($notaFrecuencia) { $eval = "Error Frecuencia: $notaFrecuencia" }
                else { $eval = "OK (Contexto)" } # Fila sana pero reportada por contexto

                $informeFinal += [PSCustomObject]@{
                    "vCenter"       = $nombreVcenter
                    "JobId"         = $fila.JobId
                    "Type"          = $fila.Type
                    "Status"        = $fila.Status
                    "StartTime"     = $fila.StartTime
                    "Location"      = $fila.Location
                    "Evaluación"    = $eval
                }
            }
        }
    }

    # 4. Definir Resultado Checklist
    $resultadoChecklist = "Resultado Esperado"
    $detalleChecklist = "Los backups se ejecutan diariamente, son programados y finalizan exitosamente."
    
    $problemasTxt = @()
    if ($hayFallidos) { $problemasTxt += "fallos de ejecución" }
    if ($hayManuales) { $problemasTxt += "ejecuciones manuales" }
    if ($hayHuecos)   { $problemasTxt += "interrupciones en la frecuencia diaria" }

    if ($problemasTxt.Count -gt 0) {
        $resultadoChecklist = "No recomendado"
        $detalleChecklist = "Se detectaron: " + ($problemasTxt -join ", ") + ". Afecta a: " + ($vcentersConProblemas -join ", ")
    }

    # 5. Exportar al Anexo
    if ($informeFinal.Count -gt 0) {
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Backup VAMI"
    }

    return [PSCustomObject]@{
        ID        = "VSP-SE-02" 
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}

function vdsBackupCheck {
    # 1. Importar datos de la hoja "Distributed Switch"
    # (Asegúrate que este sea el nombre de la hoja generado por tu script de exportación)
    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "vDS"
    } | Where-Object { -not [string]::IsNullOrEmpty($_.vCenter) }

    # Ruta fija solicitada para el mensaje de éxito
    $rutaBackups = "...\devops-powershell\reportes\vds_configuration"

    # --- CASO 1: HOJA VACÍA (NO HAY vDS) ---
    if (-not $todosLosDatos -or $todosLosDatos.Count -eq 0) {
        return [PSCustomObject]@{
            ID        = "VSP-SE-01" 
            Resultado = "Resultado Esperado"
            Detalle   = "No se detectaron Distributed Switches (vDS) en la infraestructura para respaldar."
        }
    }

    $informeFinal = @()
    $vdsAfectados = 0
    $erroresUnicos = @()

    # 2. Procesar cada fila
    foreach ($fila in $todosLosDatos) {
        $backupStatus = $fila."Backup File"
        
        # Buscamos si la celda del archivo contiene el texto de error que definimos antes
        if ($backupStatus -like "ERROR*") {
            $vdsAfectados++
            $mensajeError = $backupStatus -replace "ERROR: ", "" # Limpiamos para el reporte
            
            if ($mensajeError -notin $erroresUnicos) {
                $erroresUnicos += $mensajeError
            }

            # Agregamos al Anexo Técnico
            $informeFinal += [PSCustomObject]@{
                "vCenter"  = $fila.vCenter
                "vDS Name" = $fila.Name
                "Estado"   = "Fallo de Backup"
                "Detalle"  = $mensajeError
            }
        }
    }

    # 3. Definir Resultado del Checklist
    if ($vdsAfectados -gt 0) {
        $resultadoChecklist = "No recomendado"
        
        # Construimos un resumen de los errores
        $resumenErrores = if ($erroresUnicos.Count -gt 1) { "Múltiples errores (ver anexo)" } else { $erroresUnicos[0] }
        $detalleChecklist = "Falló el backup de $vdsAfectados vDS. Error: $resumenErrores"
        
        # 4. Exportar al Anexo (Solo si hubo fallos)
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Backup vDS"
    }
    else {
        # Caso Éxito Total
        $resultadoChecklist = "Resultado Esperado"
        $detalleChecklist = "Se realizaron los backups correctamente. Recuerde subir los archivos. Chequear path: $rutaBackups"
    }

    # 5. Retornar objeto para el Checklist
    return [PSCustomObject]@{
        ID        = "VSP-SE-01" 
        Resultado = $resultadoChecklist
        Detalle   = $detalleChecklist
    }
}



function EjecutarYActualizarChecklist {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock[]]$Tareas
    )

    Copy-Item -Path $rutaPlantilla -Destination $rutaSalidaChecklist -Force

    $resultadosChecklist = @()
    foreach ($tarea in $Tareas) {
        $resultadosChecklist += & $tarea
    }

    $excelChecklist = Open-ExcelPackage -Path $rutaSalidaChecklist

    try {
        $hojaChecklist = $excelChecklist.Workbook.Worksheets["Checklist"]
        
        foreach ($resultado in $resultadosChecklist) {
            $celdaChequeo = $hojaChecklist.Cells["A:A"] | Where-Object { $_.Text -eq $resultado.ID }
            
            if ($celdaChequeo) {
                $fila = $celdaChequeo.Start.Row
                $hojaChecklist.Cells["G$fila"].Value = $resultado.Resultado 
                $hojaChecklist.Cells["H$fila"].Value = $resultado.Detalle
                Write-Host "Actualizando checklist para '$($resultado.ID)': $($resultado.Resultado)" -ForegroundColor Green
            }

        }
    } finally {
        Close-ExcelPackage $excelChecklist
    }
}

Write-Host "Seleccione las tareas a ejecutar:"
Write-Host "1. Tareas mensuales"
Write-Host "2. Tareas trimestrales"
Write-Host "3. Tareas semestrales"
Write-Host "4. Salir"

$opcion = Read-Host "Ingrese el numero de la tarea (separe multiples opciones con comas, ej: 1,2)"
$tareasSeleccionadas = $opcion -split ','

$tareasAejecutar = [System.Collections.Generic.List[scriptblock]]::new()

[scriptblock[]]$tareasMensuales = @(
    { AnalizarSize }, { Particiones }, { SyslogCheck }, { Multipath },
    { ConsVer }, { ConsRec }, { PlacaRed }, { TSMCheck },
    { pManagement }, { vmtools }, { isos }, { placasDeRed },
    { snapshotsCheck }, { alarmCheck}
)
[scriptblock[]]$tareasTrimestrales = @(
    { endOfSupport }, { compatComponentes }, { ntpCheck }, { Drivers },
    { DNSConfig }, { Licencia }, { NIOC }, { vss }, { vCenterRoot}, 
    { vCenterCert }, { esxiCert }, { performanceHealthCheck }, { vdsBackupCheck }
)
[scriptblock[]]$tareasSemestrales = @(
    { backupCheck }, { vdsBackupCheck }
)

foreach ($tarea in $tareasSeleccionadas) {
    switch ($tarea.Trim()) {
        '1' { $tareasAejecutar.AddRange($tareasMensuales) }
        '2' { $tareasAejecutar.AddRange($tareasTrimestrales) }
        '3' { $tareasAejecutar.AddRange($tareasSemestrales) }
        '4' { Write-Host "Saliendo del script."; exit }
        default { Write-Host "Opcion no válida: '$($tarea.Trim())'" }
    }
}

if ($tareasAejecutar.Count -gt 0) {
    Write-Host "Ejecutando las tareas seleccionadas y generando la checklist..." -ForegroundColor Green
    EjecutarYActualizarChecklist -Tareas $tareasAejecutar
} else {
    Write-Host "No se seleccionó ninguna tarea válida para ejecutar." -ForegroundColor Yellow
}

Write-Host "Proceso completado. Se ha creado un nuevo archivo Excel (Anexo) en: $archivoSalida"
Write-Host "Proceso completado. Se ha creado un nuevo archivo Excel (Checklist) en: $rutaSalidaChecklist"