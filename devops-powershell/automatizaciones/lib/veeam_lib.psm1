# ==============================================================================
# LIBRERÍA DE FUNCIONES VEEAM
# Ubicación: devops-powershell/automatizaciones/lib/veeam_lib.psm1
# ==============================================================================

function Get-VeeamJobsStatus {
    <#
    .SYNOPSIS
        Obtiene el estado de los trabajos de Veeam Backup & Replication.
    .DESCRIPTION
        Recupera información sobre la última sesión de ejecución de los Jobs habilitados.
        Devuelve un objeto limpio para reporte.
    #>
    [CmdletBinding()]
    param()

    $results = @()

    try {
        # Obtenemos solo los jobs programados (ScheduleEnabled = True)
        # Esto evita ensuciar el reporte con jobs de prueba o deshabilitados.
        $jobs = Get-VBRJob | Where-Object { $_.IsScheduleEnabled -eq $true }

        foreach ($job in $jobs) {
            # FindLastSession() es el método más fiable para saber qué pasó la última vez
            $lastSession = $job.FindLastSession()

            # Normalización de valores para evitar nulos en el Excel final
            $lastResult  = if ($lastSession) { $lastSession.Result.ToString() } else { "Never Run" }
            $lastRunTime = if ($lastSession) { $lastSession.CreationTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
            $endTime     = if ($lastSession) { $lastSession.EndTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
            
            # Próxima ejecución
            $nextRun     = if ($job.Info.NextRunTime) { $job.Info.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }

            $results += [PSCustomObject]@{
                "Job Name"      = $job.Name
                "Job Type"      = $job.JobType.ToString()
                "Last Result"   = $lastResult
                "Last Run Time" = $lastRunTime
                "End Time"      = $endTime
                "Next Run Time" = $nextRun
                "Description"   = $job.Description
            }
        }
    }
    catch {
        Write-Error "Error recuperando Jobs de Veeam: $($_.Exception.Message)"
    }

    return $results
}

function Get-VeeamReposStatus {
    <#
    .SYNOPSIS
        Obtiene el estado de capacidad de los Repositorios de Veeam.
    #>
    [CmdletBinding()]
    param()

    $results = @()

    try {
        $repos = Get-VBRBackupRepository

        foreach ($repo in $repos) {
            # Cálculos de capacidad (Veeam entrega Bytes, convertimos a GB)
            $capacityGB = [math]::Round($repo.Info.Capacity / 1GB, 2)
            $freeGB     = [math]::Round($repo.Info.FreeSpace / 1GB, 2)
            
            # Cálculo de porcentaje libre (evitando división por cero)
            $freePercent = 0
            if ($capacityGB -gt 0) {
                $freePercent = [math]::Round(($freeGB / $capacityGB) * 100, 2)
            }

            # Obtener el Host asociado (si aplica)
            $hostName = "N/A"
            try { $hostObj = $repo.GetHost(); if($hostObj){ $hostName = $hostObj.Name } } catch {}

            $results += [PSCustomObject]@{
                "Repo Name"    = $repo.Name
                "Type"         = $repo.Type.ToString()
                "Path"         = $repo.Path
                "Host"         = $hostName
                "Capacity GB"  = $capacityGB
                "Free GB"      = $freeGB
                "Free %"       = $freePercent
            }
        }
    }
    catch {
        Write-Error "Error recuperando Repositorios de Veeam: $($_.Exception.Message)"
    }

    return $results
}

# Exportamos las funciones para que sean visibles al importar el módulo
Export-ModuleMember -Function Get-VeeamJobsStatus, Get-VeeamReposStatus