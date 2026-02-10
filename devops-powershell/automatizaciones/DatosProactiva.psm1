using module ".\lib\proactivas.psm1"

$credenciales = "cHJvYWN0aXZhOlBhc3N3b3JkMTIzJA=="

function Get-DatosProactivas($vcenters) {
    Start-DatosProactivas($vcenters)

    <#
	.Synopsis
    Recolección de Datos para Proactiva
	.Component
	vcenter
    .Role
    ui
	#>
}

function Start-DatosProactivas($vcenters){
    $proactiva = New-Object Proactiva
    $vCenterData = @()

    foreach($vcenter in $vcenters.conn){
        $proactiva.setCurrentVcenter($vcenter)
        Write-Host "`nProcessing vCenter: $($vcenter.Name)"
        
        Write-Host "`tGathering top-level objects..."
        $hosts = Get-VMHost -Server $vcenter
        $clusters = Get-Cluster -Server $vcenter
        $vdswitches = Get-VDSwitch -Server $vcenter
        $vms = Get-VM -Server $vcenter  # <--- LÍNEA AGREGADA

        # LÍNEA MODIFICADA para incluir el conteo de VMs
        Write-Host "`tFound $($hosts.Count) Hosts, $($clusters.Count) Clusters, $($vms.Count) VMs, $($vdswitches.Count) vDS."

        $allVms = [System.Collections.ArrayList]@()
        $allSnapshots = [System.Collections.ArrayList]@()
        
        Write-Host "`tStarting batch collection (this may take a while)..."
        for ($i = 0; $i -lt $hosts.Count; $i++) {
            $h = $hosts[$i]
            Write-Progress -Activity "Collecting data from $($vcenter.Name)" -Status "Processing Host $($i+1)/$($hosts.Count): $($h.Name)" -PercentComplete (($i / $hosts.Count) * 100)
            
            $vmsOnHost = @(Get-VM -Location $h)
            
            if ($vmsOnHost) { 
                [void]$allVms.AddRange($vmsOnHost) 
                
                $snapshotsOnHost = @(Get-Snapshot -VM $vmsOnHost)
                if ($snapshotsOnHost) { 
                    [void]$allSnapshots.AddRange($snapshotsOnHost) 
                }
            }
        }
        Write-Progress -Activity "Collection Complete" -Completed

        Write-Host "`tBatch collection finished. Processing reports..."
        $proactiva.processAlarmCheck($hosts, $vcenter)
        $proactiva.processVcenterSizing($allVms, $hosts) 
        $proactiva.processBackupActivity()
        $proactiva.processPerformanceHealth($clusters)
        $proactiva.processCertificates()
        $proactiva.processEsxi($hosts)
        $proactiva.processVcenterHealthAndInfo($allVms, $vcenter) 
        #$proactiva.processNic($hosts, $vdswitches)
        $proactiva.processVm($allVms, $clusters) 
        $proactiva.processDatastore($hosts)
        $proactiva.processSwitch($hosts)
        $proactiva.processKernelAdapters($hosts)
        $proactiva.processSnapshot($allSnapshots) 
        $proactiva.processPartitions($allVms) 
        $proactiva.processvDS($vdswitches)
        $proactiva.processLicense()
    }
    
    $file = [PSCustomObject] @{
        Result="OK"; Name="Proactiva"; Version = $global:APP_VERSION; DateTime= (Get-Date -Format "yyyy-MM-dd HH:mm");
        LocalHost= [system.environment]::MachineName; User = whoami; Endpoint=$vcenters.host; Component="vcenter";
        Report = $proactiva.getReport(); IdAutomatizacion=$credenciales;
    }

    # --- [BLOQUE CORREGIDO] ---
    # 1. Se construye la ruta completa del archivo JSON y se guarda en una variable.
    $nombreJsonGenerado = (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $file.Name + ".json"
    $rutaJsonGenerado = Join-Path -Path $global:CONFIG.REPORTS_FOLDER -ChildPath $nombreJsonGenerado

    # 2. Se guarda el archivo JSON usando esa ruta.
    $file | ConvertTo-Json -Depth 99 | Set-Content -Path $rutaJsonGenerado

    # 3. Se usa la variable global definida en portable.ps1 para escribir en resultado.txt.
    #    Añadimos el nombre del archivo que acabamos de crear a la lista.
    Add-Content -Path $global:RutaArchivoResultado -Value $nombreJsonGenerado
    # --- FIN DEL BLOQUE ---
}