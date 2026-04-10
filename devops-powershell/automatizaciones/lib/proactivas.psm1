. .\automatizaciones\lib\check-hcl.ps1
. .\automatizaciones\lib\create-alarm.ps1
. .\automatizaciones\lib\check-vsan-hcl.ps1

class Proactiva {

    [PSCustomObject[]] $toolsRefJson = (Get-Content -Path (".\automatizaciones\lib\data\vmtools-ref.json") | ConvertFrom-Json)
    [PSCustomObject[]] $vCenterSizing = (Get-Content -Path (".\automatizaciones\lib\data\vcenter-sizing.json") | ConvertFrom-Json)
    [PSCustomObject[]] $vsanhcl = (Get-Content -Path (".\automatizaciones\lib\data\vsan-hcl.json") | ConvertFrom-Json)

    [String[]] $annotations = @("VMware vCenter Server Appliance")
    
    [PSCustomObject] $VMHostNetworkAdapters = [PSCustomObject] @{}
    [PSCustomObject] $VMHostStandardSwitches = [PSCustomObject] @{}
    [PSCustomObject] $toolsReference = [PSCustomObject] @{}

    [PSCustomObject] $ioHclRef = $null
    [String] $currentVCenter = "No vCenter"

    $esxiReport = [System.Collections.ArrayList]@()
    $nicReport = [System.Collections.ArrayList]@()
    $vmReport = [System.Collections.ArrayList]@()
    $datastoreReport = [System.Collections.ArrayList]@()
    $switchReport = [System.Collections.ArrayList]@()
    $kernelAdaptersReport = [System.Collections.ArrayList]@()
    $snapshotReport = [System.Collections.ArrayList]@()
    $partitionReport = [System.Collections.ArrayList]@()
    $sizingReport = [System.Collections.ArrayList]@()
    $vdsReport = [System.Collections.ArrayList]@()
    $vNetworkReport = [System.Collections.ArrayList]@()
    $vLicenseReport = [System.Collections.ArrayList]@()
    $vCenterReport = [System.Collections.ArrayList]@()
    $certificateReport = [System.Collections.ArrayList]@()
    $performanceHealthReport = [System.Collections.ArrayList]@()
    $backupActivityReport = [System.Collections.ArrayList]@()
    $alarmCheckReport = [System.Collections.ArrayList]@()


    processEsxi($hosts) {
        Write-Host "`tProcessing ESXi..." -NoNewline
        
        # [CONTEXTO] Capturamos el servidor actual para usarlo en comandos críticos
        $serverContext = $this.currentVCenter

        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            try {
                $hclResult = $h | Check-HCL
            }
            catch {
                $hclResult = [PSCustomObject]@{
                    Model               = "No data available"
                    Supported           = $false
                    SupportedReleases   = "No data available"
                    Reference           = "No data available"
                    Note                = "Host cannot be processed"
                }
            }
            
            if (@("Connected", "Maintenance") -contains $h.ConnectionState) {
                $advSett = $h | Get-AdvancedSetting

                $advSettingShellTimeout = $advSett | Where-Object { $_.name -eq 'UserVars.ESXIShellTimeOut' }
                $ESXIShellTimeOut = if ($advSettingShellTimeout) { $advSettingShellTimeout.Value } else { $null }

                $advSettingInteractive = $advSett | Where-Object { $_.name -eq 'UserVars.ESXIShellinteractiveTimeOut' }
                $ESXIShellinteractiveTimeOut = if ($advSettingInteractive) { $advSettingInteractive.Value } else { $null }

                $advSettingLogDir = $advSett | Where-Object { $_.name -eq 'Syslog.global.logDir' }
                $SyslogGlobalLogDir = if ($advSettingLogDir) { $advSettingLogDir.Value } else { $null }

                $advSettingLogHost = $advSett | Where-Object { $_.name -eq 'Syslog.global.logHost' }
                $SyslogGlobalLogHost = if ($advSettingLogHost) { $advSettingLogHost.Value } else { $null }
                
                $ntpService = $h | Get-VMHostService | Where-Object { $_.key -eq "ntpd" }
                $ntpdRunning = if ($ntpService) { $ntpService.Running.ToString() } else { "Not Available" }

                $ntpServers = "" 
                try {
                    $ntpServers = ($h | Get-VMHostNtpServer -ErrorAction Stop) -join ","
                }
                catch {
                    $ntpServers = "" 
                }

                $certValidTo = $null
                $certStatus = "Unknown"
                $certIssuer = ""
                
                try {
                    # --- [CORRECCIÓN] Usamos -Server para asegurar que buscamos en el vCenter correcto ---
                    # También validamos que el host tenga un CertificateManager antes de consultar
                    if ($h.ExtensionData.ConfigManager.CertificateManager) {
                        $certMgr = Get-View -Id $h.ExtensionData.ConfigManager.CertificateManager -Server $serverContext -ErrorAction Stop
                        
                        if ($certMgr.CertificateInfo) {
                            $certInfo = $certMgr.CertificateInfo
                            
                            $certValidToDate = $certInfo.NotAfter
                            $certValidTo = Get-Date $certValidToDate -Format "yyyy-MM-dd"
                            $certIssuer = $certInfo.Issuer
                            
                            $hoy = Get-Date
                            if ($hoy -gt $certValidToDate) {
                                $certStatus = "Expired"
                            } elseif ($hoy.AddMonths(1) -gt $certValidToDate) {
                                $certStatus = "Expiration imminent"
                            } elseif ($hoy.AddMonths(2) -gt $certValidToDate) {
                                $certStatus = "Expiration shortly"
                            } elseif ($hoy.AddMonths(8) -gt $certValidToDate) {
                                $certStatus = "Expiration soon"
                            } else {
                                $certStatus = "Valid"
                            }
                        }
                    }
                }
                catch {
                    # Opcional: Descomentar para debug si sigue fallando
                    # Write-Warning "Fallo certificado host $($h.Name): $($_.Exception.Message)"
                }

                $hostModel = $hclResult.Model
                $supported = if ($hclResult.Supported) { "True" } else { "False" }
                $supportedReleases = $hclResult.SupportedReleases -join ","
                $reference = $hclResult.Reference
                $note = $hclResult.Note
            }
            else {
                $ESXIShellTimeOut = $null
                $ESXIShellinteractiveTimeOut = $null
                $SyslogGlobalLogDir = $null
                $SyslogGlobalLogHost = $null
                $ntpdRunning = $null
                $ntpServers = $null
                
                $certValidTo = $null
                $certStatus = "Unknown"
                $certIssuer = ""

                $hostModel = $hclResult.Model
                $supported = if ($hclResult.Supported) { "True" } else { "False" }
                $supportedReleases = $hclResult.SupportedReleases -join ","
                $reference = $hclResult.Reference
                $note = $hclResult.Note
            }

            $this.esxiReport += [PSCustomObject] @{
                vCenter                     = $this.currentVCenter # Usamos el nombre del objeto
                Hostname                    = $h.Name
                Model                       = $hostModel
                Datacenter                  = ($h | Get-Datacenter).Name
                Cluster                     = ($h | Get-Cluster).Name
                ESXiVersion                 = $h.ExtensionData.Summary.Config.Product.FullName
                ConnectionState             = $h.ConnectionState.ToString()
                MemoryGB                    = [math]::Round($h.MemoryTotalGB, 3)
                CpuModel                    = $h.ProcessorType
                CpuSpeed                    = $h.ExtensionData.summary.hardware.CpuMhz
                DnsServer                   = $h.ExtensionData.Config.Network.DnsConfig.Address -join ","
                NtpServer                   = $ntpServers
                NtpdRunning                 = $ntpdRunning
                "Cert Valid To"             = $certValidTo
                "Cert Status"               = $certStatus
                "Cert Issuer"               = $certIssuer
                ESXIShellTimeOut            = $ESXIShellTimeOut
                ESXIShellinteractiveTimeOut = $ESXIShellinteractiveTimeOut
                PowerManagement             = $h.ExtensionData.Hardware.CpuPowerManagementInfo.CurrentPolicy
                SyslogGlobalLogDir          = $SyslogGlobalLogDir
                SyslogGlobalLogHost         = $SyslogGlobalLogHost
                Supported                   = $supported
                "Supported Releases"        = $supportedReleases
                Note                        = $note
                Reference                   = $reference
            }
        }
    }
    
    processNic($hosts, $vdswitches) {
        Write-Host "`tProcessing ESXi IO..." -NoNewline
        $vmHosts = $hosts | Where-Object { @("Connected", "Maintenance") -contains $_.ConnectionState }
        $devices = $vmHosts | Get-VMHostPciDevice | Where-Object { $_.DeviceClass -eq "MassStorageController" -or $_.DeviceClass -eq "NetworkController" -or $_.DeviceClass -eq "SerialBusController" } 
        for ($count = 0; $count -lt $devices.length; $count++) {
            Show-Progress $devices.length ($count + 1)
            $device = $devices[$count]
            if ($device.DeviceName -like "*USB*" -or $device.DeviceName -like "*iLO*" -or $device.DeviceName -like "*iDRAC*") {
                continue
            }
            
            $vid = [String]::Format("{0:x4}", $device.VendorId)
            $did = [String]::Format("{0:x4}", $device.DeviceId)
            $svid = [String]::Format("{0:x4}", $device.SubVendorId)
            $ssid = [String]::Format("{0:x4}", $device.SubDeviceId)
            $hclReference = $this.getHCLReference()
            foreach ($entry in $hclReference.data.ioDevices) {
                if (($vid -eq $entry.vid) -and ($did -eq $entry.did) -and ($svid -eq $entry.svid) -and ($ssid -eq $entry.ssid)) {
                    $isNic = $device.DeviceClass -eq "NetworkController"
                    $isHba = $device.DeviceClass -eq "MassStorageController" -or $device.DeviceClass -eq "SerialBusController"
                    $deviceData = $null
                    $vmnicdetail = $null
                    $standardSwitches = $null
                    $version = $null
                    $vsanCompatibility = $null
                    $esxicli = $device.VMHost | Get-EsxCli -V2
                    if ($isNic) {
                        $standardSwitches = $this.getHostStandardSwitches($device.VMHost)
                        $deviceData = $esxicli.network.nic.list.Invoke() | Where-Object { $_.PCIDevice -like '*' + $device.Id }
                        
                        if ($deviceData) {
                            $vmnicdetail = $esxicli.network.nic.get.Invoke(@{nicname = $deviceData.Name })
                            $version = $vmnicdetail.DriverInfo.Version
                            $vsanCompatibility = $null
                            
                            $this.nicReport += [PSCustomObject] @{
                                "vCenter"                   = $this.currentVCenter;
                                "Hostname"                  = $device.VMHost.name;
                                "Placa"                     = $deviceData.Name;
                                "Controlador"               = $device.DeviceName;
                                "Vendor"                    = $device.VendorName;
                                "Driver"                    = $vmnicdetail.DriverInfo.Driver;
                                "Version"                   = $version;
                                "Firmware"                  = $vmnicdetail.DriverInfo.FirmwareVersion;
                                "Vid"                       = $vid; "Did" = $did; "Svid" = $svid; "ssid" = $ssid;
                                "ESXi Release"              = $device.VMHost.ExtensionData.Summary.Config.Product.FullName;
                                #"ESXi Supported Releases"   = $entry.releases -join ",";
                                "URL"                       = $entry.url;
                                #"vSAN Compatibility"        = $vsanCompatibility -join ",";
                                #"Switch"                    = $this.getSwitchNameForPNic($deviceData, $vdswitches, $standardSwitches);
                            }
                        }
                    }
                    elseif ($isHba) {
                        $deviceData = $device.VMHost | Get-VMHostHba -ErrorAction SilentlyContinue | Where-Object { $_.PCI -like '*' + $device.Id } 
                        
                        if ($deviceData) {
                            $vibname = $deviceData.Driver -replace "_", "-"
                            $version = ($esxicli.software.vib.list.invoke() | Where-Object { $_.Name -match "^(scsi-|sata-|)$vibname" }).Version -join ", "
                            $vsanCompatibility = Get-VsanHclDatabase $vid $did $svid $ssid $($this.vsanhcl)
                            
                            $this.nicReport += [PSCustomObject] @{
                                "vCenter"                   = $this.currentVCenter;
                                "Hostname"                  = $device.VMHost.name;
                                "Placa"                     = $deviceData.Device;
                                "Controlador"               = $device.DeviceName;
                                "Vendor"                    = $device.VendorName;
                                "Driver"                    = $deviceData.Driver;
                                "Version"                   = $version;
                                "Firmware"                  = $null;
                                "Vid"                       = $vid; "Did" = $did; "Svid" = $svid; "ssid" = $ssid;
                                "ESXi Release"              = $device.VMHost.ExtensionData.Summary.Config.Product.FullName;
                                "ESXi Supported Releases"   = $entry.releases -join ",";
                                "URL"                       = $entry.url;
                                "vSAN Compatibility"        = $vsanCompatibility -join ",";
                                "Switch"                    = $null;
                            }
                        }
                    }
                    break;
                }
            }
        }
    }
    processVm($vms, $clusters) {
        $vms = $vms | ForEach-Object { $_ }

        Write-Host "`tProcessing network adapters..." -NoNewline
        $snapshots = $vms | Get-Snapshot
        $connectedIsos = $vms | Get-CDDrive | Where-Object { $null -ne $_.IsoPath }
        $allNetworkAdapters = $vms | Get-NetworkAdapter
        $this.processvNetwork($allNetworkAdapters)
        Write-Host "`tProcessing VMs..." -NoNewline
        for ($count = 0; $count -lt $vms.Count; $count++) { # Usamos .Count que es más robusto
            $vm = $vms[$count]
            Show-Progress $vms.Count ($count + 1)
            $networkAdapters = $allNetworkAdapters | Where-Object { $_.ParentId -eq $vm.Id }
            $nicCount = $networkAdapters.length
            $toolsRequredVersion = $this.getToolsReference($vm.VMHost)
            $this.vmReport += [PSCustomObject]@{
                vCenter              = $this.currentVCenter
                VM                   = $vm.Name
                Cluster              = ($clusters | Where-Object { $_.ExtensionData.Host -contains $vm.VMHost.Id }).name
                Host                 = $vm.VMHost.name
                ConnectionState      = $vm.ExtensionData.Runtime.ConnectionState.ToString()
                State                = $vm.PowerState.ToString()
                vCPU                 = $vm.NumCpu
                "Memory MB"          = $vm.MemoryMB
                HardwareVersion      = $vm.HardwareVersion
                Snapshots            = ($snapshots | Where-Object { $_.VM -eq $vm }).length
                ToolsStatus          = try { $vm.ExtensionData.Guest.ToolsStatus.ToString() } Catch { "VM has not been scanned" };
                ToolsVersion         = $vm.ExtensionData.Guest.ToolsVersion
                ToolsRequiredVersion = $toolsRequredVersion
                "SO (vCenter)"       = $vm.ExtensionData.Config.GuestFullName
                "SO (Tools)"         = $vm.ExtensionData.Guest.GuestFullName
                IsoConnected         = ($connectedIsos | Where-Object { $_.ParentId -eq $vm.Id }).IsoPath
                Adapter_01           = if ($nicCount -gt 0) { $networkAdapters[0].Type.ToString() } else { "" } 
                Adapter_02           = if ($nicCount -gt 1) { $networkAdapters[1].Type.ToString() } else { "" } 
                Adapter_03           = if ($nicCount -gt 2) { $networkAdapters[2].Type.ToString() } else { "" } 
                Adapter_04           = if ($nicCount -gt 3) { $networkAdapters[3].Type.ToString() } else { "" } 
                Adapter_05           = if ($nicCount -gt 4) { $networkAdapters[4].Type.ToString() } else { "" } 
                Adapter_06           = if ($nicCount -gt 5) { $networkAdapters[5].Type.ToString() } else { "" } 
                Adapter_07           = if ($nicCount -gt 6) { $networkAdapters[6].Type.ToString() } else { "" } 
                Adapter_08           = if ($nicCount -gt 7) { $networkAdapters[7].Type.ToString() } else { "" } 
                Adapter_09           = if ($nicCount -gt 8) { $networkAdapters[8].Type.ToString() } else { "" } 
                Adapter_10           = if ($nicCount -gt 9) { $networkAdapters[9].Type.ToString() } else { "" } 
            }
        }
    }

    processvNetwork($allNetworkAdapters) {
        for ($count = 0; $count -lt $allNetworkAdapters.length; $count++) {
            Show-Progress $allNetworkAdapters.length ($count + 1)
            $networkadapter = $allNetworkAdapters[$count]
            $this.vNetworkReport += [PSCustomObject]@{
                vCenter         = $this.currentVCenter
                VM              = $networkadapter.Parent.Name
                Cluster         = ($networkadapter.Parent.VMHost | Get-Cluster).Name
                Host            = $networkadapter.Parent.VMHost.Name
                Status          = $networkadapter.Parent.PowerState
                Mac             = $networkadapter.MacAddress
                Connected       = if ($networkadapter.ConnectionState.Connected) { "True" } else { "False" }
                StartsConnected = if ($networkadapter.ConnectionState.StartConnected) { "True" } else { "False" }
            }
        }
    }
    processLicense() {
        Write-Host "`tProcessing licenses..."
        foreach ($licenseManager in (Get-View LicenseManager)) {
            #-Server $vCenter.Name
            foreach ($license in $licenseManager.Licenses) {
                $licenseProp = $license.Properties
                $licenseExpiryInfo = $licenseProp | Where-Object { $_.Key -eq 'expirationDate' } | Select-Object -ExpandProperty Value
                if ($license.Name -eq 'Product Evaluation') {
                    $expirationDate = 'Evaluation'
                } #if ($license.Name -eq 'Product Evaluation')
                elseif ($null -eq $licenseExpiryInfo) {
                    $expirationDate = 'Never'
                } #elseif ($null -eq $licenseExpiryInfo)
                else {
                    $expirationDate = $licenseExpiryInfo
                } #else #if ($license.Name -eq 'Product Evaluation')
    
                if ($license.Total -eq 0) {
                    $totalLicenses = 'Unlimited'
                } #if ($license.Total -eq 0)
                else {
                    $totalLicenses = $license.Total
                } #else #if ($license.Total -eq 0)

                $productName = $licenseProp | Where-Object { $_.Key -eq 'ProductName' } | Select-Object -ExpandProperty Value
                $productVersion = $licenseProp | Where-Object { $_.Key -eq 'ProductVersion' } | Select-Object -ExpandProperty Value
    
                $this.vLicenseReport += [PSCustomObject]@{
                    Name           = $license.Name
                    LicenseKey     = $license.LicenseKey
                    ExpirationDate = $expirationDate
                    ProductName    = if ($null -eq ($productName)) { "No product name" } else { $productName } 
                    ProductVersion = if ($null -eq ($productName)) { "No product version" } else { $productVersion }
                    EditionKey     = $license.EditionKey
                    Total          = $totalLicenses
                    Used           = $license.Used
                    CostUnit       = $license.CostUnit
                    vCenter        = $this.currentVCenter
                }
            } #foreach ($license in $licenseManager.Licenses)
        }
    }
    
    processDatastore($hosts) {
        Write-Host "`tProcessing Datastores..." -NoNewline

        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $esxName = $h.Name

            $allDatastores = $h | Get-Datastore
            
            $vmfsDatastores = $allDatastores | Where-Object { $_.Type -ne "vsan" }
            $esx = Get-View -ViewType HostSystem -Property Name, Config.StorageDevice -Filter @{"Name" = "^$esxName" }
            
            foreach ($lun in $esx.Config.StorageDevice.MultipathInfo.Lun) {
                $scsiLun = $esx.Config.StorageDevice.ScsiLun | Where-Object { $_.Key -eq $lun.Lun }
                $canon = $scsiLun.CanonicalName
                $datastore = ($vmfsDatastores | Where-Object { ($_.extensiondata.info.vmfs.extent | Select-Object -expand diskname) -like $canon }).name

                if ($null -ne $datastore) {
                    $policy = if ($lun.Policy.Policy -match "_FIXED") { "Fixed" }
                    elseif ($lun.Policy.Policy -match "_MRU") { "MostRecentlyUsed" }
                    elseif ($lun.Policy.Policy -match "_RR") { "RoundRobin" }
                    else { "Unknown" }

                    $this.datastoreReport += [PSCustomObject] @{
                        vCenter   = $this.currentVCenter
                        Hostname  = $esxName
                        Datastore = $datastore
                        Policy    = $policy
                    }
                }
            }

            # Logica para los vSAN, dado que estos no tienen LUNs ni politicas de Multipath esta parte es un poco irrelevante
            $vsanDatastores = $allDatastores | Where-Object { $_.Type -eq "vsan" }
            foreach ($vsanDs in $vsanDatastores) {
                if (-not ($this.datastoreReport | Where-Object { $_.Datastore -eq $vsanDs.Name })) {
                    $this.datastoreReport += [PSCustomObject] @{
                        vCenter   = $this.currentVCenter
                        Hostname  = $esxName 
                        Datastore = $vsanDs.Name
                        Policy    = "vSAN" 
                    }
                }
            }
        }
    }
    
    processSwitch($hosts) {
        Write-Host "`tProcessing Standard Switches..." -NoNewline
        $totalPG = 0
        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $cluster = ($h | Get-Cluster).Name
            foreach ($sw in ($h | Get-VirtualSwitch -Standard)) {
                $portGroups = $sw | Get-VirtualPortGroup
                $totalPG += $portGroups.length
                foreach ($pg in $portGroups) {
                    if ($null -ne $pg.vLanId) {
                        $this.switchReport += [PSCustomObject] @{
                            vCenter   = $this.currentVCenter
                            ESXi      = $h.name
                            Cluster   = $cluster
                            PortGroup = $pg.Name
                            Switch    = $sw.Name
                            vLAN      = $pg.vLanId
                        }
                    }
                }
            }
        }
    }
    processSnapshot($snapshots) {
        $snapshots = $snapshots | ForEach-Object { $_ }

        Write-Host "`tProcessing Snapshots...";
        for ($count = 0; $count -lt $snapshots.Count; $count++) { # Usamos .Count
            Show-Progress $snapshots.Count ($count + 1)
            $s = $snapshots[$count]
            $this.snapshotReport += [PSCustomObject] @{
                vCenter  = $this.currentVCenter
                VM       = $s.VM.Name
                Snapshot = $s.Name
                Fecha    = ($s.Created | Get-Date -Format "yyyy-MM-dd HH:mm")
                SizeMB   = [int]$s.SizeMB
            }
        }
    }
    processPartitions($vms) {
        $vms = $vms | ForEach-Object { $_ }
        
        Write-Host "`tProcessing Partitions..." -NoNewline
        for ($count = 0; $count -lt $vms.Count; $count++) { # Usamos .Count
            $vm = $vms[$count];
            Show-Progress $vms.Count ($count + 1)
            if ($vm.ExtensionData.Config.Annotation -in $this.annotations) {
                foreach ($partition in $vm.ExtensionData.Guest.Disk) {
                    $freePercentage = [math]::Round(($partition.FreeSpace / $partition.Capacity) * 100, 2);
                    $this.partitionReport += [PSCustomObject]@{
                        vCenter    = $this.currentVCenter
                        VM         = $vm.name
                        Annotation = $vm.ExtensionData.Config.Annotation
                        Disk       = $partition.DiskPath
                        "Free %"   = $freePercentage
                    }
                }
            }
        }
    }
    processKernelAdapters($hosts) {
        Write-Host "`tProcessing VMkernel Adapters..." -NoNewline
        $kernelAdapters = $hosts | Get-VMHostNetworkAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "vmk[0-9]+" }
        for ($count = 0; $count -lt $kernelAdapters.length; $count++) {
            Show-Progress $kernelAdapters.length ($count + 1)
            $ka = $kernelAdapters[$count]
            $this.kernelAdaptersReport += [PSCustomObject] @{
                Host       = $ka.VMHost.Name
                Name       = $ka.Name
                IP         = $ka.IP
                MTU        = $ka.MTU
                PortGroup  = $ka.PortGroupName 
                Management = $ka.ManagementTrafficEnabled
                vMotion    = $ka.VMotionEnabled
            }
        }
    }
    
    processVcenterSizing($vms, $hosts) {
        Write-Host "`tProcessing Sizing..." -NoNewline
        
        # --- Extracción robusta del nombre del vCenter ---
        $rawVCenter = $this.currentVCenter
        $vcenterName = if ($rawVCenter.Name) { $rawVCenter.Name } else { $rawVCenter }
        
        if ([string]::IsNullOrEmpty($vcenterName)) {
            Write-Warning " -> Error: No se pudo determinar el nombre del vCenter."
            return
        }
        
        # 1. BÚSQUEDA DE VM (Lógica Robusta)
        $vcenterShortName = ($vcenterName).Split('.')[0]
        
        $targetVM = $vms | Where-Object { 
            $_.Name -eq $vcenterName -or 
            $_.Name -eq $vcenterShortName 
        } | Select-Object -First 1

        # Fallback: Intentar por Annotation
        if (-not $targetVM) {
            $targetVM = $vms | Where-Object { $_.ExtensionData.Config.Annotation -in $this.annotations } | Select-Object -First 1
        }

        # Si no aparece, reportamos "No detectado"
        if (-not $targetVM) {
            $this.sizingReport += [PSCustomObject]@{
                vCenter             = $vcenterName
                VM                  = "No encontrada"
                Annotation          = "-"
                vCPU                = "-"
                "Memory GB"         = "-"
                "Cantidad de VMs"   = $vms.Count
                "Cantidad de Hosts" = $hosts.Count
                "Sizing actual"     = "Unknown (VM not found)"
            }
            return
        }

        # 2. MATCHING DE SIZING
        $sizingFound = $false
        
        for ($i = $this.vCenterSizing.vsphere.Count - 1; $i -ge 0; $i--) {
            $ref = $this.vCenterSizing.vsphere[$i]
            
            # --- [CORRECCIÓN] Volvemos a la lógica original para el nombre ---
            # Usamos la propiedad .ToString tal cual estaba en tu código viejo
            $refName = $ref.ToString
            
            # (Fallback de seguridad por si acaso la propiedad no existe, para no dejar vacío)
            if (-not $refName) { $refName = if ($ref.size) { $ref.size } else { $ref.name } }

            $refCpu = $ref.vcpus
            $refRam = $ref.ram

            # Comparación (Mayor o Igual)
            if (($targetVM.NumCpu -ge $refCpu) -and ($targetVM.MemoryGB -ge $refRam)) {
                
                $this.sizingReport += [PSCustomObject]@{
                    vCenter             = $vcenterName
                    VM                  = $targetVM.Name
                    Annotation          = if ($targetVM.ExtensionData.Config.Annotation) { $targetVM.ExtensionData.Config.Annotation } else { "-" }
                    vCPU                = $targetVM.NumCpu
                    "Memory GB"         = [math]::Round($targetVM.MemoryGB, 0)
                    "Cantidad de VMs"   = $vms.Count
                    "Cantidad de Hosts" = $hosts.Count
                    "Sizing actual"     = $refName
                }
                
                $sizingFound = $true
                break # Encontrado el perfil más alto que cumple, salimos.
            }
        }

        # 3. CASO CUSTOM
        if (-not $sizingFound) {
            $this.sizingReport += [PSCustomObject]@{
                vCenter             = $vcenterName
                VM                  = $targetVM.Name
                Annotation          = if ($targetVM.ExtensionData.Config.Annotation) { $targetVM.ExtensionData.Config.Annotation } else { "-" }
                vCPU                = $targetVM.NumCpu
                "Memory GB"         = [math]::Round($targetVM.MemoryGB, 0)
                "Cantidad de VMs"   = $vms.Count
                "Cantidad de Hosts" = $hosts.Count
                "Sizing actual"     = "Custom / Undefined"
            }
        }
        
        Write-Host " -> OK." -ForegroundColor Green
    }

    processvDS($vdswitches) {
        Write-Host "`tProcessing vDS (and Backup)..." -NoNewline;
        
        $backupPath = Join-Path -Path $global:CONFIG.REPORTS_FOLDER -ChildPath "vds_configuration"

        # 2. Verificamos si existe. Si no, la creamos.
        if (-not (Test-Path $backupPath)) {
            try {
                New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
            }
            catch {
                Write-Warning "No se pudo crear la carpeta de backup en '$backupPath'. Los respaldos fallarán."
            }
        }

        foreach ($vds in $vdswitches) {
            # 3. Construimos un nombre de archivo ordenado (Resuelve tu TODO)
            # Formato: vCenter_NombreSwitch_Fecha.zip
            $vCenterClean = $this.currentVCenter -replace '[^a-zA-Z0-9]', '_' # Limpiamos caracteres raros
            $dateStr = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
            $fileName = "${vCenterClean}_$($vds.Name)_$dateStr.zip"
            $fullFilePath = Join-Path -Path $backupPath -ChildPath $fileName

            # 4. Ejecutamos el backup
            try {
                Export-VDSwitch -VDSwitch $vds -Destination $fullFilePath -ErrorAction Stop
            }
            catch {
                Write-Warning "Error al exportar backup del vDS '$($vds.Name)': $($_.Exception.Message)"
                $fileName = "ERROR: $($_.Exception.Message)"
            }

            # 5. Lógica de reporte original
            $niocEnabled = if ($vds.ExtensionData.Config.NetworkResourceManagementEnabled) { "True" } else { "False" }
            
            $this.vdsReport += [PSCustomObject]@{
                vCenter        = $this.currentVCenter # Agregué esta columna para consistencia
                Name           = $vds.Name
                MTU            = $vds.Mtu
                "NIOC Enabled" = $niocEnabled
                "Backup File"  = $fileName # Opcional: Agregamos el nombre del backup al reporte
            }
        }
        Write-Host " -> OK." -ForegroundColor Green
    }
    

    processAlarmCheck($hosts, $vcenterConnection) {
        Write-Host "`tProcessing Alarm Check (Datastore Target & Dynamic Script)..." -NoNewline
        
        $serverContext = $this.currentVCenter
        $vcenterName = $vcenterConnection.Name
        
        $uniqueId = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $alarmName = "Falso Positivo DS $uniqueId"
        $sourceAlarmName = "N/A"
        $scriptPath = $null
        $reportResult = "Pendiente"

        # 1. SELECCIONAR OBJETIVO (DATASTORE)
        $targetDatastore = Get-Datastore -Server $serverContext | Where-Object { $_.State -eq "Available" -and $_.Accessible } | Select-Object -First 1
        
        if (-not $targetDatastore) {
            Write-Warning " -> No hay datastores accesibles."
            $this.alarmCheckReport += [PSCustomObject]@{
                vCenter = $vcenterName; Host = "N/A"; "Alarm Name" = "N/A"; "Script Path" = "N/A"; Result = "No hay datastores accesibles"; Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
            return
        }
        $targetName = $targetDatastore.Name 

        # 2. BUSCAR ALARMA FUENTE (Dinámicamente)
        # Obtenemos todas las alarmas y buscamos la primera que tenga un script configurado
        $allAlarms = Get-AlarmDefinition -Server $serverContext -ErrorAction SilentlyContinue
        
        foreach ($def in $allAlarms) {
            try {
                $info = $def.ExtensionData.Info
                if ($info.Action -and $info.Action.Action) {
                    foreach ($triggerAction in $info.Action.Action) {
                        if ($triggerAction.Action -is [VMware.Vim.RunScriptAction]) {
                            $potentialScript = $triggerAction.Action.Script
                            if (-not [string]::IsNullOrWhiteSpace($potentialScript)) {
                                $scriptPath = $potentialScript
                                $sourceAlarmName = $def.Name
                                break # Encontramos el script en esta acción, salimos del bucle interior
                            }
                        }
                    }
                }
            } catch {}
            
            if ($scriptPath) { break } # Ya encontramos un script válido, salimos del bucle principal
        }

        # Validación por si el vCenter no tiene NINGUNA alarma con script
        if ([string]::IsNullOrEmpty($scriptPath)) {
            $this.alarmCheckReport += [PSCustomObject]@{ 
                vCenter = $vcenterName; Host = $targetName; "Alarm Name" = "N/A"; "Script Path" = "N/A"; Result = "Ninguna alarma en vCenter tiene script configurado"; Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") 
            }
            return
        }

        # 3. PREPARAR MÉTRICA (disk.used.latest)
        $metricId = $null
        try {
            $perfMgr = Get-View (Get-View ServiceInstance -Server $serverContext).Content.PerfManager -Server $serverContext
            $counterInfo = $perfMgr.PerfCounter | Where-Object { $_.GroupInfo.Key -eq "disk" -and $_.NameInfo.Key -eq "used" -and $_.RollupType -eq "latest" } | Select-Object -First 1
            if ($counterInfo) { $metricId = [int]$counterInfo.Key } else { throw "Metric missing" }
        }
        catch {
            $this.alarmCheckReport += [PSCustomObject]@{ vCenter = $vcenterName; Host = $targetName; "Alarm Name" = $sourceAlarmName; "Script Path" = $scriptPath; Result = "ERROR API Métrica"; Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
            return
        }

        # 4. CREAR Y DISPARAR ALARMA
        try {
            # A. Limpieza preventiva
            $existing = Get-AlarmDefinition -Name $alarmName -Entity $targetDatastore -Server $serverContext -ErrorAction SilentlyContinue
            if ($existing) { Remove-AlarmDefinition $existing -Server $serverContext -Confirm:$false }

            # B. Definición Base
            $spec = New-Object VMware.Vim.AlarmSpec
            $spec.Name = $alarmName
            $spec.Description = "DevOps Smoke Test"
            $spec.Enabled = $true 
            
            # C. Expresión: Metric Alarm envuelta en OR
            $expression = New-Object VMware.Vim.MetricAlarmExpression
            $expression.Operator = "isAbove"
            $expression.Type = "Datastore"
            $expression.Yellow = 10 # 10 KB (Siempre se cumple)
            $expression.Red = 20    # 20 KB (Siempre se cumple)
            
            $expression.Metric = New-Object VMware.Vim.PerfMetricId
            $expression.Metric.CounterId = $metricId
            $expression.Metric.Instance = ""
            
            $spec.Expression = New-Object VMware.Vim.OrAlarmExpression
            $spec.Expression.Expression += $expression

            # D. Acción (Estructura Compleja)
            $actionGroup = New-Object VMware.Vim.GroupAlarmAction
            $actionTrigger = New-Object VMware.Vim.AlarmTriggeringAction
            
            $scriptAction = New-Object VMware.Vim.RunScriptAction
            $scriptAction.Script = $scriptPath
            $actionTrigger.Action = $scriptAction
            
            $actionTrigger.Green2yellow = $true
            $actionTrigger.Yellow2red = $true
            
            $trans1 = New-Object VMware.Vim.AlarmTriggeringActionTransitionSpec
            $trans1.StartState = [VMware.Vim.ManagedEntityStatus]::green
            $trans1.FinalState = [VMware.Vim.ManagedEntityStatus]::yellow
            $trans1.Repeats = $false
            $actionTrigger.TransitionSpecs += $trans1
            
            $trans2 = New-Object VMware.Vim.AlarmTriggeringActionTransitionSpec
            $trans2.StartState = [VMware.Vim.ManagedEntityStatus]::yellow
            $trans2.FinalState = [VMware.Vim.ManagedEntityStatus]::red
            $trans2.Repeats = $false
            $actionTrigger.TransitionSpecs += $trans2

            $actionGroup.Action += $actionTrigger
            $spec.Action = $actionGroup

            # Setting
            $spec.Setting = New-Object VMware.Vim.AlarmSetting
            $spec.Setting.ReportingFrequency = 0
            $spec.Setting.ToleranceRange = 0
            $spec.ActionFrequency = 0

            # E. Crear
            $alarmManager = Get-View AlarmManager -Server $serverContext
            $moref = $alarmManager.CreateAlarm($targetDatastore.ExtensionData.MoRef, $spec)
            
            # F. Esperar y Validar
            Start-Sleep -Seconds 15
            
            $events = Get-VIEvent -Entity $targetDatastore -MaxSamples 50 -Server $serverContext | Where-Object { $_.FullFormattedMessage -like "*$alarmName*" }
            
            if ($events | Where-Object { $_ -is [VMware.Vim.AlarmScriptCompleteEvent] }) {
                $reportResult = "SUCCESS"
            } elseif ($events | Where-Object { $_ -is [VMware.Vim.AlarmScriptFailedEvent] }) {
                $reportResult = "FAILED: Script execution error"
            } else {
                # Fallback Visual
                $dsView = Get-View $targetDatastore.Id -Property TriggeredAlarmState -Server $serverContext
                $triggeredState = $dsView.TriggeredAlarmState | Where-Object { $_.Alarm.Value -eq $moref.Value }
                if ($triggeredState -and $triggeredState.OverallStatus -eq "red") {
                    $reportResult = "SUCCESS (Falso positivo ejecutado)"
                } else {
                    $reportResult = "WARNING: Alarm created but trigger failed"
                }
            }

            # G. Limpieza
            $created = Get-View $moref -Server $serverContext
            $created.RemoveAlarm()
        }
        catch {
            Write-Warning "`nError en prueba de alarma: $($_.Exception.Message)"
            $reportResult = "ERROR: $($_.Exception.Message)"
            $al = Get-AlarmDefinition -Name $alarmName -Entity $targetDatastore -Server $serverContext -ErrorAction SilentlyContinue
            if ($al) { Remove-AlarmDefinition $al -Server $serverContext -Confirm:$false }
        }

        # [REPORTE FINAL]
        $this.alarmCheckReport += [PSCustomObject]@{
            vCenter         = $vcenterName
            Host            = $targetName
            "Alarm Name"    = $sourceAlarmName
            "Script Path"   = $scriptPath
            Result          = $reportResult
            Timestamp       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        
        $this.auditAlarmConfiguration($serverContext, $vcenterName)
        Write-Host " -> OK." -ForegroundColor Green
    }

    # ---------------------------------------------------------
    # FUNCIÓN AUXILIAR: AUDITORÍA DE LISTA DE ALARMAS
    # ---------------------------------------------------------
    auditAlarmConfiguration($serverContext, $vcenterName) {
        
        # 1. Detectar si hay vSAN habilitado en algún cluster
        $isVsanEnabled = $false
        try {
            $clusters = Get-Cluster -Server $serverContext -ErrorAction SilentlyContinue
            foreach ($c in $clusters) {
                if ($c.VsanEnabled) { $isVsanEnabled = $true; break }
            }
        } catch {}

        # 2. Definir las listas de alarmas
        $standardAlarms = @(
            "Host battery status", "Host connection failure", "Host hardware fan status", 
            "Host hardware power status", "Host hardware system board status", "Host hardware temperature status", 
            "Host hardware voltage", "Network connectivity lost", "Network uplink redundancy degraded", 
            "Network uplink redundancy lost", "Unmanaged workload detected on SIOC-enabled datastore", 
            "Cannot connect to storage", "Certificate Status", "ESXi Host Certificate Status", 
            "Insufficient vSphere HA failover resources", "vSphere HA failover in progress"
        )

        $vsanAlarms = @(
            "Errors occurred on the disk(s) of a vSAN host", "vSAN capacity utilization alarm 'What if the most consumed host fails'",
            "vSAN capacity utilization alarm 'Storage space'", "vSAN cluster alarm 'vSAN daemon liveness'",
            "vSAN network alarm 'Hosts with connectivity issues'", "vSAN network alarm 'Network latency check'",
            "vSAN network alarm 'vSAN cluster partition'", "vSAN online health alarm 'vSAN critical alert regarding a potential data inconsistency'",
            "vSAN physical disk alarm 'Component limit'", "vSAN physical disk alarm 'Congestion'",
            "vSAN physical disk alarm 'Operation'", "vSAN physical disk alarm 'Disk capacity'",
            "vSAN stretched cluster alarm 'Witness host not found'", "vSAN stretched cluster alarm for 'Site latency' health",
            "vSAN performance service alarm 'Verbose mode'", "vSAN performance service alarm 'Network diagnostic mode'"
        )

        # 3. Combinar listas según corresponda
        $alarmsToCheck = $standardAlarms
        if ($isVsanEnabled) {
            $alarmsToCheck += $vsanAlarms
        }

        # 4. Iterar y Verificar
        foreach ($alarmName in $alarmsToCheck) {
            $status = "Not Found"
            $path = "N/A"
            
            # Buscamos la definición exacta
            $def = Get-AlarmDefinition -Name $alarmName -Server $serverContext -ErrorAction SilentlyContinue | Select-Object -First 1
            
            if ($def) {
                $status = "No Script Configured"
                
                # Inspeccionamos las acciones
                try {
                    $info = $def.ExtensionData.Info
                    if ($info.Action -and $info.Action.Action) {
                        foreach ($triggerAction in $info.Action.Action) {
                            if ($triggerAction.Action -is [VMware.Vim.RunScriptAction]) {
                                $script = $triggerAction.Action.Script
                                if (-not [string]::IsNullOrWhiteSpace($script)) {
                                    $path = $script
                                    $status = "OK"
                                } else {
                                    $status = "Script Action Empty"
                                }
                                break 
                            }
                        }
                    }
                } catch {
                    $status = "Error reading actions"
                }
            }

            # Guardamos en el MISMO reporte
            $this.alarmCheckReport += [PSCustomObject]@{
                vCenter      = $vcenterName
                Host         = "-" # No aplica a host específico
                "Alarm Name" = $alarmName
                "Script Path"= $path
                Result       = $status
                Timestamp    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
        }
    }
    
    processVcenterHealthAndInfo($allVms, $vcenterConnection) {
        Write-Host "`tProcessing vCenter General & Health Info..." -NoNewline
        
        # --- PARTE 1: Datos Generales ---
        $vcenterName = $vcenterConnection.Name
        $version     = $vcenterConnection.Version
        $build       = $vcenterConnection.Build
        
        $vcenterShortName = ($vcenterName).Split('.')[0]
        $vcenterVmObj = $allVms | Where-Object { $_.Name -eq $vcenterName -or $_.Name -eq $vcenterShortName } | Select-Object -First 1
        $vmName = if ($vcenterVmObj) { $vcenterVmObj.Name } else { "No Encontrada (Físico/Otro)" }

        # --- PARTE 2: Datos Root (CIS) ---
        $rootUser = "root"
        $expirationDate = "N/A"
        $daysRemaining = "N/A"
        $rootStatus = "Unknown"

        Import-Module VMware.VimAutomation.Cis.Core -ErrorAction SilentlyContinue

        try {
            # Obtenemos el servicio de cuentas locales
            $accountService = Get-CisService -Name "com.vmware.appliance.local_accounts" -ErrorAction Stop | Select-Object -First 1
            
            if ($accountService) {
                $rootInfo = $accountService.get("root")
                $hoy = Get-Date

                # LÓGICA CORREGIDA: Verificamos si existe el dato de la fecha directamente
                if ($rootInfo.password_expires_at) {
                    $expirationDateObj = Get-Date $rootInfo.password_expires_at
                    $expirationDate = $expirationDateObj.ToString("yyyy-MM-dd HH:mm")
                    
                    $diasRestantes = ($expirationDateObj - $hoy).Days
                    $daysRemaining = $diasRestantes
                    
                    # Definición de Estado
                    if ($diasRestantes -lt 0) { $rootStatus = "Expirada" }
                    elseif ($diasRestantes -lt 30) { $rootStatus = "Expira Pronto (Crítico)" }
                    else { $rootStatus = "Válida" }
                }
                else {
                    # Si el campo fecha viene vacío, es que no expira
                    $expirationDate = "Nunca"
                    $rootStatus = "Válida (Sin expiración)"
                }
            }
            else {
                $rootStatus = "Error: Servicio LocalAccounts no disponible"
            }
        }
        catch {
            # Captura de error silenciosa pero efectiva para el reporte
            $rootStatus = "Error API: $($_.Exception.Message)"
        }

        # --- PARTE 3: Reporte ---
        $this.vCenterReport += [PSCustomObject]@{
            "vCenter Server"  = $vcenterName
            "VM Name"         = $vmName
            "Version"         = $version
            "Build"           = $build
            "Root User"       = $rootUser
            "Expiration Date" = $expirationDate
            "Days Remaining"  = $daysRemaining
        }
        
        Write-Host " -> OK." -ForegroundColor Green
    }
    
    processCertificates() {
        Write-Host "`tProcessing Certificates (API Method)..." -NoNewline
        
        Import-Module VMware.VimAutomation.Cis.Core -ErrorAction SilentlyContinue
        
        $hoy = (Get-Date)
        $fechaLimite = $hoy.AddDays(30)
        
        try {
            # Seleccionamos el PRIMERO (-First 1) para asegurar que es un objeto único y no un array
            $tlsService = Get-CisService -Name "com.vmware.vcenter.certificate_management.vcenter.tls" | Select-Object -First 1
            
            if ($tlsService) {
                $tlsCertData = $tlsService.get()

                $validTo = Get-Date $tlsCertData.valid_to
                $validFrom = Get-Date $tlsCertData.valid_from

                $status = "Valid"
                if ($hoy -gt $validTo) { $status = "Expirado" }
                elseif ($fechaLimite -gt $validTo) { $status = "Expira Pronto" }

                $this.certificateReport += [PSCustomObject]@{
                    vCenter      = $this.currentVCenter
                    Ubicacion    = "Machine SSL"
                    Subject      = $tlsCertData.subject_dn
                    Status       = $status
                    "Valid From" = $validFrom
                    "Valid Until"= $validTo
                    Emisor       = $tlsCertData.issuer_dn
                }
            }
        }
        catch {
            Write-Warning "Error obteniendo Machine SSL: $($_.Exception.Message)"
        }

        try {
            $signingCertService = Get-CisService -Name "com.vmware.vcenter.certificate_management.vcenter.signing_certificate" | Select-Object -First 1
            
            if ($signingCertService) {
                $signingCertsData = $signingCertService.get().signing_cert_chains.cert_chain

                foreach ($pemString in $signingCertsData) {
                    try {
                        # Limpieza manual (Old School) para compatibilidad
                        $cleanBase64 = $pemString -replace "-----BEGIN CERTIFICATE-----", ""
                        $cleanBase64 = $cleanBase64 -replace "-----END CERTIFICATE-----", ""
                        $cleanBase64 = $cleanBase64 -replace "\s", "" # Quita espacios y saltos de linea

                        $certBytes = [System.Convert]::FromBase64String($cleanBase64)
                        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(,$certBytes)

                        $certType = "VMCA_ROOT"
                        if ($cert.Subject -match "CN=ssoserverSign") { $certType = "STS" }

                        $status = "Valid"
                        if ($hoy -gt $cert.NotAfter) { $status = "Expirado" }
                        elseif ($fechaLimite -gt $cert.NotAfter) { $status = "Expira Pronto" }

                        $this.certificateReport += [PSCustomObject]@{
                            vCenter      = $this.currentVCenter
                            Ubicacion    = $certType
                            Subject      = $cert.Subject
                            Status       = $status
                            "Valid From" = $cert.NotBefore
                            "Valid Until"= $cert.NotAfter
                            Emisor       = $cert.Issuer
                        }
                    } catch {}
                }
            }
        }
        catch {
            Write-Warning "Error obteniendo lista de VMCA/STS: $($_.Exception.Message)"
        }

        # --- C. TRUSTED ROOT CHAINS ---
        try {
            $rootService = Get-CisService -Name "com.vmware.vcenter.certificate_management.vcenter.trusted_root_chains" | Select-Object -First 1
            
            if ($rootService) {
                $chains = $rootService.list().chain

                foreach ($chainId in $chains) {
                    $certChainData = $rootService.get($chainId)
                    $rawCertData = $certChainData.cert_chain.cert_chain

                    if ($rawCertData -is [Array]) { $pemString = $rawCertData -join "`n" }
                    else { $pemString = $rawCertData }

                    $pattern = "(?ms)-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----"
                    $matches = [regex]::Matches($pemString, $pattern)

                    foreach ($match in $matches) {
                        try {
                            $singleCertPem = $match.Value
                            $cleanBase64 = $singleCertPem -replace "-----BEGIN CERTIFICATE-----", ""
                            $cleanBase64 = $cleanBase64 -replace "-----END CERTIFICATE-----", ""
                            $cleanBase64 = $cleanBase64 -replace "\s", "" 
                            
                            $certBytes = [System.Convert]::FromBase64String($cleanBase64)
                            $rootCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(,$certBytes)

                            $status = "Valid"
                            if ($hoy -gt $rootCert.NotAfter) { $status = "Expirado" }
                            elseif ($fechaLimite -gt $rootCert.NotAfter) { $status = "Expira Pronto" }

                            $this.certificateReport += [PSCustomObject]@{
                                vCenter      = $this.currentVCenter
                                Ubicacion    = "Trusted Root"
                                Subject      = $rootCert.Subject
                                Status       = $status
                                "Valid From" = $rootCert.NotBefore
                                "Valid Until"= $rootCert.NotAfter
                                Emisor       = $rootCert.Issuer
                            }
                        } catch {}
                    }
                }
            }
        }
        catch {
            Write-Warning "Error obteniendo lista de Trusted Roots: $($_.Exception.Message)"
        }
        
        Write-Host " -> OK." -ForegroundColor Green
    }

    processPerformanceHealth($clusters) {
    
        Write-Host "`tProcessing Performance Health Check..." -NoNewline
    
        # Obtener el vCenter actual para el reporte
        $vCenterName = $this.currentVCenter
        
        # Métricas clave a verificar (Promedio de uso)
        $metricsToCheck = @(
            "cpu.usage.average",
            "mem.usage.average"
        )

        # Métricas adicionales para Hosts (donde realmente residen los datos)
        $hostMetricsToCheck = @(
            "net.usage.average",
            "disk.usage.average"
        )

        # Fecha de inicio: últimas 2 horas
        $start = (Get-Date).AddHours(-2)

        if (-not $clusters) {
             Write-Host " -> Sin Clusters para evaluar." -ForegroundColor Yellow
             return
        }

        foreach ($cluster in $clusters) {
            # Inicializamos estados en "Sin Datos"
            $cpuStatus = "Sin Datos"
            $memStatus = "Sin Datos"
            $netStatus = "Sin Datos"
            $diskStatus = "Sin Datos"
            $dbStatus = "Posible Error DB" 

            # --- PARTE 1: Verificar CPU/Memoria (Nivel Cluster) ---
            try {
                # Solicitamos solo CPU y Memoria al CLUSTER
                $statsCluster = Get-Stat -Entity $cluster -Stat $metricsToCheck -Start $start -MaxSamples 1 -ErrorAction SilentlyContinue

                if ($statsCluster) {
                    if ($statsCluster | Where-Object { $_.MetricId -eq "cpu.usage.average" }) { $cpuStatus = "OK" }
                    if ($statsCluster | Where-Object { $_.MetricId -eq "mem.usage.average" }) { $memStatus = "OK" }

                    # Si CPU y Memoria tienen datos, el servicio de stats funciona
                    if ($cpuStatus -eq "OK" -and $memStatus -eq "OK") {
                        $dbStatus = "OK"
                    }
                }
            }
            catch {
                $dbStatus = "Error de Consulta (Cluster)"
            }

            # --- PARTE 2: Verificar Red/Disco (Nivel Host) ---
            # Si el servicio de estadísticas del Cluster está OK, verificamos métricas de Host
            if ($dbStatus -eq "OK") {
                # Obtenemos los hosts del cluster. Usamos Get-View para que sea más rápido.
                $hosts = Get-View -ViewType HostSystem -Property Name -SearchRoot $cluster.Id 

                if ($hosts) {
                    # Intentamos obtener una muestra de las métricas de Red/Disco de UN host.
                    # Con que un host devuelva datos es suficiente para verificar la salud del servicio.
                    try {
                        $statsHost = Get-Stat -Entity $hosts[0].Name -Stat $hostMetricsToCheck -Start $start -MaxSamples 1 -ErrorAction Stop

                        if ($statsHost) {
                            # Verificamos Red y Disco en la muestra del Host
                            if ($statsHost | Where-Object { $_.MetricId -eq "net.usage.average" }) { $netStatus = "OK" }
                            if ($statsHost | Where-Object { $_.MetricId -eq "disk.usage.average" }) { $diskStatus = "OK" }
                        }
                    }
                    catch {
                        # Si falla, el DBStatus sigue siendo OK, pero las métricas son "Sin Datos"
                        Write-Warning "Error consultando métricas de Host para $($cluster.Name). Error: $($_.Exception.Message)"
                    }
                }
            }

            # Guardamos el reporte simplificado
            $this.performanceHealthReport += [PSCustomObject]@{
                vCenter      = $vCenterName 
                Cluster      = $cluster.Name
                "Health DB"  = $dbStatus
                "CPU Stats"  = $cpuStatus
                "Mem Stats"  = $memStatus
                "Net Stats"  = $netStatus
                "Disk Stats" = $diskStatus
            }
        }
        Write-Host " -> OK." -ForegroundColor Green
    }

    processBackupActivity() {
        Write-Host "`tProcessing Backup Activity (Smart Fallback + Context)..." -NoNewline
        
        # 1. Obtener nombre limpio del vCenter para el reporte
        $rawVCenter = $this.currentVCenter
        $cisFQDN = if ($rawVCenter.Name) { $rawVCenter.Name } else { $rawVCenter }
        
        Import-Module VMware.VimAutomation.Cis.Core -ErrorAction SilentlyContinue

        # 2. Obtener la conexión CIS específica para este vCenter
        $cisConnection = $global:cisConnections[$cisFQDN]

        # Fallback: Si no está en el llavero, buscar en las conexiones globales por nombre
        if (-not $cisConnection) {
            $cisConnection = $global:DefaultCisServers | Where-Object { $_.Name -eq $cisFQDN } | Select-Object -First 1
        }

        if (-not $cisConnection) {
            Write-Warning " -> No hay conexión CIS activa para $cisFQDN. (Error de permisos/conexión previa)"
            $this.backupActivityReport += [PSCustomObject]@{
                vCenter = $cisFQDN; Type = "-"; Status = "Error de Conexión CIS"; "Data Transfer" = "-"; Location = "-"; StartTime = "-"; EndTime = "-"; Duration = "-"
            }
            return
        }

        try {
            $allJobs = @()
            $usandoDetalles = $false

            # --- INTENTO 1: Servicio de DETALLES (Rico en datos) ---
            try {
                $detailServiceName = "com.vmware.appliance.recovery.backup.job.details"
                # [CORRECCIÓN] Usamos -Server (no -CisServer)
                $detailsService = Get-CisService -Name $detailServiceName -Server $cisConnection -ErrorAction SilentlyContinue | Select-Object -First 1

                if ($detailsService) {
                    $jobsMap = $detailsService.list($null)
                    
                    if ($jobsMap) {
                        foreach ($entry in $jobsMap.GetEnumerator()) {
                            $jobObj = $entry.Value
                            $jobIdKey = $entry.Key
                            if ($null -eq $jobObj.id) {
                                $jobObj | Add-Member -MemberType NoteProperty -Name "id" -Value $jobIdKey -Force
                            }
                            $allJobs += $jobObj
                        }
                        $usandoDetalles = $true
                    }
                }
            } catch { }

            # --- INTENTO 2: Servicio SIMPLE (Solo estado y fechas) ---
            if ($allJobs.Count -eq 0) {
                $simpleServiceName = "com.vmware.appliance.recovery.backup.job"
                # [CORRECCIÓN] Usamos -Server
                $simpleService = Get-CisService -Name $simpleServiceName -Server $cisConnection -ErrorAction SilentlyContinue | Select-Object -First 1
                
                if ($simpleService) {
                    # Obtenemos la lista y aseguramos que sea un array de strings limpio
                    $rawList = $simpleService.list()
                    # A veces viene como objeto con propiedad Value, a veces directo. Esto lo normaliza:
                    if ($rawList.Value) { $jobIds = $rawList.Value } else { $jobIds = $rawList }
                    
                    foreach ($jid in $jobIds) {
                         try { 
                             $j = $simpleService.get($jid)
                             if (!$j.id) { $j | Add-Member -Name "id" -Value $jid -Force }
                             $allJobs += $j
                         } catch {}
                    }
                }
            }

            # --- Generación del Reporte ---
            if ($allJobs.Count -eq 0) {
                 # Caso VACÍO: Marcamos como "No Configurado"
                 $this.backupActivityReport += [PSCustomObject]@{
                    vCenter          = $cisFQDN
                    Type             = "-"
                    Status           = "No Configurado"
                    "Data Transfer"  = "-"
                    Location         = "-"
                    StartTime        = "-"
                    EndTime          = "-"
                    Duration         = "-"
                 }
                 Write-Host " -> Sin historial (Marcado como No Configurado)." -ForegroundColor Yellow
                 return
            }
            
            # Procesamiento Normal
            $backupHistory = $allJobs | Sort-Object start_time -Descending | Select-Object -First 7
            
            foreach ($job in $backupHistory) {
                $status = if ($job.state) { $job.state } else { $job.status }
                $startTime = Get-Date $job.start_time -Format "yyyy-MM-dd HH:mm:ss"
                $endTime = Get-Date $job.end_time -Format "yyyy-MM-dd HH:mm:ss"
                
                $duration = "N/A"
                if ($job.end_time -and $job.start_time) {
                    $ts = New-TimeSpan -Start $job.start_time -End $job.end_time
                    $duration = "{0:hh\:mm\:ss}" -f $ts
                }

                $location = "N/A"; $sizeGB = "N/A"; $type = "N/A"

                if ($usandoDetalles) {
                    if ($job.location) { $location = $job.location }
                    if ($job.type) { $type = $job.type.ToString() }
                    if ($job.size) { 
                        $sizeGB = [math]::Round($job.size / 1GB, 2).ToString() + " GB"
                    }
                }

                $this.backupActivityReport += [PSCustomObject]@{
                    vCenter          = $cisFQDN
                    Type             = $type
                    Status           = $status
                    "Data Transfer"  = $sizeGB
                    Location         = $location
                    StartTime        = $startTime
                    EndTime          = $endTime
                    Duration         = $duration
                }
            }
            
            Write-Host " -> OK (Procesados $($backupHistory.Count) trabajos)." -ForegroundColor Green
        }
        catch {
            Write-Warning "`nError en Backup Activity: $($_.Exception.Message)"
            $this.backupActivityReport += [PSCustomObject]@{
                vCenter = $cisFQDN; Type = "-"; Status = "Error API: $($_.Exception.Message)"; "Data Transfer" = "-"; Location = "-"; StartTime = "-"; EndTime = "-"; Duration = "-"
            }
        }
    }

    setCurrentVcenter($vcenter) {
        $this.currentVCenter = $vcenter
    }

    [PSCustomObject] getReport() {
        return [PSCustomObject] @{
            "vCenter"           = $this.vCenterReport;
            "ESXi"              = $this.esxiReport;
            "ESXi IO"           = $this.nicReport;
            "VM"                = $this.vmReport;
            "Datastores"        = $this.datastoreReport;
            "Standard Switch"   = $this.switchReport;
            "VMkernel Adapters" = $This.kernelAdaptersReport;
            "Snapshot"          = $this.snapshotReport;
            "Partitions"        = $this.partitionReport;
            "Sizing"            = $this.sizingReport;
            "vDS"               = $this.vdsReport;
            "vNetwork"          = $this.vNetworkReport;
            "vLicense"          = $this.vLicenseReport;
            "Certficate"        = $this.certificateReport;
            "PerformanceHealth" = $this.performanceHealthReport;
            "BackupActivity"    = $this.backupActivityReport;
            "Falso Positivo"    = $this.alarmCheckReport          
        }
    }

    [PSCustomObject] getHCLReference() {    
        if ($null -eq $this.ioHclRef) {
            $this.ioHclRef = Get-Content -Path ($global:CONFIG.PLUGINS_FOLDER + "\lib\data\vmware-iohcl.json")
            $this.ioHclRef = $this.ioHclRef | ConvertFrom-Json
            return $this.ioHclRef
        }
        else {
            return $this.ioHclRef
        }
    }

    [Array] getHostNetworkAdaptersForVDS($vds) {
        if ($null -eq ($this.VMHostNetworkAdapters | Get-Member -Name $vds.name)) {
            $this.VMHostNetworkAdapters | Add-Member -NotePropertyName $vds.name -NotePropertyValue (Get-VMHostNetworkAdapter -DistributedSwitch $vds)
        }
        return $this.VMHostNetworkAdapters.($vds.name)
    }

    [String] getSwitchNameForPNic($pnic, $vdswitches, $standardSwitches) {
        foreach ($vdswitch in $vdswitches) {
            $networkAdapters = $this.getHostNetworkAdaptersForVDS($vdswitch)
            $thisAdapter = $networkAdapters | Where-Object { $_.mac -eq $pnic.MACAddress }
            if ($null -ne $thisAdapter) {
                return $vdswitch.name
            }
        }
        foreach ($sSwitch in $standardSwitches) {
            if ($sSwitch.nic -contains $pnic.name) {
                return $sSwitch.name
            }
        }
        return "None"
    }

    [Array] getHostStandardSwitches($vmhost) {
        if ($null -eq ($this.VMHostStandardSwitches | Get-Member -Name $vmhost.name)) {
            $this.VMHostStandardSwitches | Add-Member -NotePropertyName $vmhost.name -NotePropertyValue ($vmhost | Get-VirtualSwitch -Standard)
        }
        return $this.VMHostStandardSwitches.($vmhost.name)
    }

    [String] getToolsReference($esxi) {
        $toolsbuild = $this.toolsReference | Get-Member -Name $esxi.name
        if ($null -eq $toolsbuild) {
            $this.toolsReference | Add-Member -NotePropertyName $esxi.name -NotePropertyValue ($this.toolsRefJson | Where-Object { $_.esxiBuild -eq $esxi.ExtensionData.Config.Product.build }).toolsBuild
        }
        return $this.toolsReference.($esxi.name)
    }
}