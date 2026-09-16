[CmdletBinding()]
param (
    [switch]$AutoDismountISO
)

# ====================================================================
# VMware Daily Morning Health Check - Enterprise Master v24.0
# (Features: Integrated GitHub Logic from v12.2, Factor-Centric Consolidated Layout,
#            Datastore Used-Bar Render Fix, Multi-vCenter Merged Tables)
# ====================================================================

$ErrorActionPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

Import-Module VMware.PowerCLI -ErrorAction SilentlyContinue
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCEIP $false -Scope Session -Confirm:$false -WarningAction SilentlyContinue | Out-Null

# ------------------------------------------------------------
# 1. 환경, 계정 및 자격 증명 설정
# ------------------------------------------------------------
$SecConfigFile = "$PSScriptRoot\sec_config.clixml"

if (Test-Path $SecConfigFile) {
    $secConfig = Import-Clixml -Path $SecConfigFile
    $vcAccounts  = $secConfig.VcAccounts
    $SmtpUser    = $secConfig.SmtpUser
    $BstrPass    = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secConfig.SmtpPass)
    $SmtpPass    = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BstrPass)
    $BstrToken   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secConfig.GitHubToken)
    $GitHubToken = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BstrToken)
    $GitHubRepo  = $secConfig.GitHubRepo
} else {
    $vcAccounts = @{
        "vcsa01.rangers.lab"       = @{ User = "administrator@vsphere.local"; Pass = "#" }
        "kr-vcs9-mgmt.rangers.lab" = @{ User = "administrator@vsphere.local"; Pass = "#" }
    }
    $SmtpUser   = "#@gmail.com"
    $SmtpPass   = "#"
    $GitHubToken = "#"
    $GitHubRepo  = "#"
}


$SmtpServer    = "smtp.gmail.com"
$SmtpPort      = 587
$BccRecipients = @(
    "#",
    "#"
)

# 기본 실행 변수
$From       = $SmtpUser
$UtcToday   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd')
$Subject    = "🛡️ VMware Infrastructure Master O&M Health Check ($UtcToday UTC)"
$StartTime  = (Get-Date).AddHours(-24).ToUniversalTime()

$HistoryFilePath = "$PSScriptRoot\vm_inventory_history.json"
$LogFilePath     = "$PSScriptRoot\HealthCheck_Execution.log"

# 디테일 로거 함수
function Write-DetailedLog {
    param([string]$Message, [string]$Level = "INFO")
    $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timeStr] [$Level] $Message"
    
    switch ($Level) {
        "ERROR"  { Write-Host $logLine -ForegroundColor Red }
        "WARN"   { Write-Host $logLine -ForegroundColor Yellow }
        "SUCCESS"{ Write-Host $logLine -ForegroundColor Green }
        Default  { Write-Host $logLine -ForegroundColor Gray }
    }
    
    $logLine | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

# 임계치 설정
$SnapAgeLimitDays      = 7
$CertWarningDays       = 30
$LatencyWarningMs      = 15
$CpuMemWarningPct      = 70
$CpuMemCriticalPct     = 85
$CpuVirtualizationPct  = 400   # CPU 가상화율 경고 임계치 (400% = 4:1)
$MemVirtualizationPct  = 120   # Memory 가상화율 경고 임계치 (120%)
$VcDiskWarningPct      = 80    # VCSA 디스크 사용률 경고 임계치 (80%)

# ------------------------------------------------------------
# 2. DNS 조회를 위한 캐시 엔진
# ------------------------------------------------------------
$dnsCache = @{}
function Get-DNSHostname ($ip) {
    if (-not $ip -or $ip -eq "127.0.0.1" -or $ip -eq "Unknown IP" -or $ip -eq "Local / Internal") { return $ip }
    if ($dnsCache.ContainsKey($ip)) { return $dnsCache[$ip] }
    
    try {
        $hostName = [System.Net.Dns]::GetHostEntry($ip).HostName
        $result = "$ip ($hostName)"
    } catch {
        $result = $ip
    }
    $dnsCache[$ip] = $result
    return $result
}

# ------------------------------------------------------------
# 3. Outlook-Safe HTML 렌더링 헬퍼 함수
# ------------------------------------------------------------
$EC = @{
    bg = "#f1f4f9"; surface = "#ffffff"; border = "#dde1e8"; 
    text = "#171923"; muted = "#64748b"; th_bg = "#0f172a"; th_text = "#ffffff";
    green = "#16a34a"; yellow = "#d97706"; red = "#dc2626"; blue = "#2563eb"; pink = "#db2777"; gray = "#94a3b8"; purple = "#7e22ce"
}

function Get-BadgeHtml {
    param($Text, $ColorStr)
    return "<span style='display:inline-block;padding:3px 8px;font-size:11px;font-weight:bold;color:#ffffff;background-color:$ColorStr;border-radius:10px;white-space:nowrap;'>$Text</span>"
}

function Get-OutlookBar {
    param([double]$pct)
    if ($pct -gt 100) { $pct = 100 }
    if ($pct -lt 0) { $pct = 0 }
    $remain = 100 - $pct
    $color = if ($pct -ge $CpuMemCriticalPct) { $EC.red } elseif ($pct -ge $CpuMemWarningPct) { $EC.yellow } else { $EC.green }
    
    $barHtml = "<table width='100px' border='0' cellpadding='0' cellspacing='0' style='border:1px solid #dde1e8;background-color:#f1f5f9;margin-top:4px;'><tr>"
    if ($pct -gt 0) { $barHtml += "<td width='$pct%' style='background-color:$color;height:6px;font-size:1px;line-height:1px;'>&nbsp;</td>" }
    if ($remain -gt 0) { $barHtml += "<td width='$remain%' style='background-color:transparent;height:6px;font-size:1px;line-height:1px;'>&nbsp;</td>" }
    $barHtml += "</tr></table>"
    return $barHtml
}

# 💡 스토리지 실제 사용량 % 전용 프로그래스 바 (사용률이 적을수록 게이지 바가 작게 출력)
function Get-StorageBar {
    param([double]$usedPct)
    if ($usedPct -gt 100) { $usedPct = 100 }
    if ($usedPct -lt 0) { $usedPct = 0 }
    $remain = 100 - $usedPct
    
    $color = if ($usedPct -ge 90) { $EC.red } elseif ($usedPct -ge 85) { $EC.yellow } else { $EC.green }
    
    $barHtml = "<table width='120px' border='0' cellpadding='0' cellspacing='0' style='border:1px solid #dde1e8;background-color:#f1f5f9;margin-top:4px;'><tr>"
    if ($usedPct -gt 0) { $barHtml += "<td width='$usedPct%' style='background-color:$color;height:8px;font-size:1px;line-height:1px;'>&nbsp;</td>" }
    if ($remain -gt 0) { $barHtml += "<td width='$remain%' style='background-color:transparent;height:6px;font-size:1px;line-height:1px;'>&nbsp;</td>" }
    $barHtml += "</tr></table>"
    return $barHtml
}

function Build-EmailTable {
    param([string[]]$Headers, $Rows, [string]$EmptyMsg = "데이터가 없습니다.")
    $thHtml = ($Headers | ForEach-Object { "<th align='left' valign='middle' style='padding:8px 10px;background-color:$($EC.th_bg);color:$($EC.th_text);font-size:12px;font-family:Arial,sans-serif;border:1px solid $($EC.border);white-space:nowrap;position:sticky;top:0;'>$_</th>" }) -join ""
    
    if (-not $Rows -or $Rows.Count -eq 0) {
        $bodyHtml = "<tr><td colspan='$($Headers.Count)' align='center' valign='middle' style='padding:15px;background-color:$($EC.surface);border:1px solid $($EC.border);color:$($EC.muted);font-size:12px;font-family:Arial,sans-serif;'>$EmptyMsg</td></tr>"
    } else {
        $bodyHtml = ($Rows | ForEach-Object {
            $tds = ($_ | ForEach-Object { "<td align='left' valign='top' style='padding:8px 10px;background-color:$($EC.surface);border:1px solid $($EC.border);font-size:12px;color:$($EC.text);font-family:Arial,sans-serif;word-break:break-all;'>$_</td>" }) -join ""
            "<tr class='data-row'>$tds</tr>"
        }) -join ""
    }
    return "<table class='report-table' width='100%' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;margin-bottom:20px;width:100%;'><thead><tr>$thHtml</tr></thead><tbody>$bodyHtml</tbody></table>"
}

function Build-SectionHeader {
    param($Icon, $Title)
    return "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='margin:30px 0 10px 0;width:100%;'><tr><td style='font-size:15px;font-weight:bold;color:$($EC.text);font-family:Arial,sans-serif;border-left:4px solid $($EC.blue);padding-left:10px;'>$Icon $Title</td></tr></table>"
}

function Get-FastSnapshotTree {
    param($tree, $vmName)
    $res = @()
    foreach ($node in $tree) {
        $res += [PSCustomObject]@{ VM = $vmName; Name = $node.Name; Created = $node.CreateTime }
        if ($node.ChildSnapshotList) { $res += Get-FastSnapshotTree -tree $node.ChildSnapshotList -vmName $vmName }
    }
    return $res
}

# ------------------------------------------------------------
# 4. 데이터 수집 및 점검 초기화
# ------------------------------------------------------------
Write-DetailedLog "==========================================" "INFO"
Write-DetailedLog "vSphere Daily Health Check v24.0 스크립트 실행 시작" "INFO"

$CategorizedAlerts = @{
    "Hardware" = @()
    "Network"  = @()
    "Resource" = @()
    "Storage"  = @()
    "Security" = @()
    "Snapshot" = @()
}

$PrevVMHistory = @{}
if (Test-Path $HistoryFilePath) {
    try { $PrevVMHistory = Get-Content $HistoryFilePath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
}
$CurrentVMHistory = @{}

# 💡 검증 팩터(Factor)별 모든 vCenter 수집 행 저장 배열
$AllVmTrendRows        = @()
$AllVcStatusRows       = @()
$AllVcPartitionRows    = @()
$AllClusterRows        = @()
$AllHostRows           = @()
$AllDatastoreRows      = @()
$AllVmRows             = @()
$AllIsoMountedRows     = @()
$AllHourlyLoginRows    = @()
$AllSourceAppRows      = @()
$AllAuthFailRows       = @()
$AllAdminActionRows    = @()
$AllOrphanedVmRows     = @()
$AllCriticalEventRows  = @()

$vMotionTotalCount = 0

foreach ($vc in $vcAccounts.Keys) {
    Write-DetailedLog "[$vc] 데이터 수집 및 O&M 분석 진행 중..." "INFO"

    $conn = $null
    try {
        $user = $vcAccounts[$vc].User
        $pass = if ($vcAccounts[$vc].Pass -is [System.Security.SecureString]) {
            [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($vcAccounts[$vc].Pass))
        } else { $vcAccounts[$vc].Pass }

        $conn = Connect-VIServer -Server $vc -User $user -Password $pass -WarningAction SilentlyContinue -ErrorAction Stop

        $global:alarmCache = @{}
        function Get-AlarmName($alarmMoRef, $connection) {
            if ($null -eq $alarmMoRef -or -not $alarmMoRef.Value) { return "Unknown Alarm" }
            if (-not $global:alarmCache.ContainsKey($alarmMoRef.Value)) {
                try {
                    $alarmView = Get-View -Id $alarmMoRef -Property Info.Name -Server $connection -ErrorAction Stop
                    if ($alarmView -and $alarmView.Info) { $global:alarmCache[$alarmMoRef.Value] = $alarmView.Info.Name }
                    else { $global:alarmCache[$alarmMoRef.Value] = "Alarm ($($alarmMoRef.Value))" }
                } catch { $global:alarmCache[$alarmMoRef.Value] = "Alarm ($($alarmMoRef.Value))" }
            }
            return $global:alarmCache[$alarmMoRef.Value]
        }

        # 데이터 일괄 추출
        $allVMs = Get-View -ViewType VirtualMachine -Property Name, Runtime.Host, Runtime.PowerState, Runtime.ConnectionState, Snapshot, Runtime.MaxCpuUsage, Summary.QuickStats, Runtime.ConsolidationNeeded, Guest.ToolsRunningStatus, Guest.ToolsVersionStatus2, Config.Template, Summary.Config.NumCpu, Summary.Config.MemorySizeMB, Summary.Storage.Committed, Summary.Storage.Uncommitted, Guest.IpAddress, Guest.Net, Config.Hardware.NumCoresPerSocket, Guest.Disk, Config.Hardware.Device -Server $conn
        $allHosts = Get-View -ViewType HostSystem -Property Name, Runtime.ConnectionState, Runtime.BootTime, Runtime.HealthSystemRuntime, Summary, Config, TriggeredAlarmState, Hardware.CpuInfo.NumCpuCores, Hardware.MemorySize, Config.DateTimeInfo, Runtime.InMaintenanceMode -Server $conn
        $allClusters = Get-Cluster -Server $conn
        $allDatastores = Get-Datastore -Server $conn
        $hostMap = @{}; foreach ($h in $allHosts) { $hostMap[$h.MoRef.Value] = $h.Name }

        $nowUtcStr = (Get-Date).ToUniversalTime().ToString('MM-dd HH:mm') + " UTC"

        # VM 증감 분석
        $currVmNames = $allVMs | Select-Object -ExpandProperty Name
        $CurrentVMHistory[$vc] = $currVmNames
        $prevVmNames = if ($PrevVMHistory.PSObject.Properties[$vc]) { $PrevVMHistory.$vc } else { @() }
        
        $newVMs = $currVmNames | Where-Object { $_ -notin $prevVmNames }
        $deletedVMs = $prevVmNames | Where-Object { $_ -notin $currVmNames }
        
        $diffCount = $currVmNames.Count - $prevVmNames.Count
        $diffStr = if ($diffCount -gt 0) { "<span style='color:$($EC.purple);font-weight:bold;'>+$diffCount 대</span>" } elseif ($diffCount -lt 0) { "<span style='color:$($EC.red);font-weight:bold;'>$diffCount 대</span>" } else { "변동 없음 (0)" }
        $newStr = if ($newVMs) { ($newVMs | ForEach-Object { (Get-BadgeHtml "+" $EC.purple) + " $_" }) -join "<br>" } else { "-" }
        $delStr = if ($deletedVMs) { ($deletedVMs | ForEach-Object { (Get-BadgeHtml "-" $EC.red) + " $_" }) -join "<br>" } else { "-" }

        $AllVmTrendRows += ,@("<b>$vc</b>", "$($prevVmNames.Count) 대", "$($currVmNames.Count) 대", $diffStr, $newStr, $delStr)

        # 0. vCenter 어플라이언스 점검
        $vcCertDays = "N/A"
        try {
            $req = [Net.HttpWebRequest]::Create("https://$vc"); $req.Timeout = 3000; $req.Method = "HEAD"; $req.GetResponse() | Out-Null
            if ($req.ServicePoint.Certificate) { $vcCertDays = ([datetime]::Parse($req.ServicePoint.Certificate.GetExpirationDateString()) - (Get-Date)).Days }
        } catch {}
        
        if ($vcCertDays -ne "N/A" -and $vcCertDays -le $CertWarningDays) { 
            $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$vc; Time=$nowUtcStr; Message="vCenter 인증서 만료 임박 (D-$vcCertDays 일)" }
        }
        $vcCertBadge = if ($vcCertDays -le $CertWarningDays -and $vcCertDays -ne "N/A") { (Get-BadgeHtml "D-$vcCertDays" $EC.red) } else { (Get-BadgeHtml "D-$vcCertDays" $EC.green) }
        
        try {
            $vcsaVM = $allVMs | Where-Object { $_.Name -match "vCenter|VCSA|vcsa|$vc" } | Select-Object -First 1
            if (-not $vcsaVM) { $vcsaVM = $allVMs | Where-Object { $null -ne $_.Guest.Disk -and $_.Guest.Disk.Count -gt 5 } | Select-Object -First 1 }

            if ($vcsaVM -and $vcsaVM.Guest.Disk) {
                foreach ($disk in $vcsaVM.Guest.Disk) {
                    $pPath = if ($disk.DiskPath) { $disk.DiskPath } elseif ($disk.Path) { $disk.Path } else { "Partition" }
                    $capGB = [math]::Round($disk.Capacity / 1GB, 1)
                    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
                    $usedGB = [math]::Round(($disk.Capacity - $disk.FreeSpace) / 1GB, 1)
                    $usedPct = if ($capGB -gt 0) { [math]::Round(($usedGB / $capGB) * 100, 1) } else { 0 }
                    
                    $pctDisplay = if ($usedPct -ge $VcDiskWarningPct) {
                        $CategorizedAlerts["Storage"] += [PSCustomObject]@{ VC=$vc; Target="$vc ($pPath)"; Time=$nowUtcStr; Message="vCenter 파티션 용량 위험 ($usedPct% 사용 중)" }
                        "<b style='color:$($EC.red);'>$usedPct% (경고)</b>"
                    } else { "$usedPct%" }
                    $AllVcPartitionRows += ,@("<b>$vc</b>", "<b>$pPath</b>", "$capGB GB", "$usedGB GB", "$freeGB GB", $pctDisplay, "ext4/xfs")
                }
            }
        } catch {}

        $si = Get-View -Id "ServiceInstance" -Server $conn
        $rootFolder = Get-View -Id $si.Content.RootFolder -Property TriggeredAlarmState -Server $conn
        $vcAlarmStr = ""
        if ($rootFolder -and $rootFolder.TriggeredAlarmState) {
            $vcAlarms = $rootFolder.TriggeredAlarmState | ForEach-Object { Get-AlarmName $_.Alarm $conn }
            $vcAlarmStr = "<span style='color:$($EC.red);font-weight:bold;'>" + ($vcAlarms -join "<br>") + "</span>"
            $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$vc; Time=$nowUtcStr; Message="vCenter 어플라이언스 알람 감지 ($($vcAlarms.Count) 건) - $($vcAlarms -join ', ')" }
        } else { $vcAlarmStr = (Get-BadgeHtml "OK" $EC.green) }
        
        $AllVcStatusRows += ,@("<b>$vc</b>", $vcCertBadge, $vcAlarmStr)

        # Cluster HA/DRS/EVC Audit
        try {
            $rawEvents = Get-VIEvent -Server $conn -Start $StartTime -MaxSamples 2000
            $vmotionEvents = $rawEvents | Where-Object { 
                $_.GetType().Name -match "VmRelocatedEvent|VmMovedEvent|VmBeingRelocatedEvent" -or
                $_.FullFormattedMessage -match "relocated|migrated"
            }
            $vMotionTotalCount += $vmotionEvents.Count
        } catch {}

        foreach ($cl in $allClusters) {
            $haStatus = if ($cl.HAEnabled) { (Get-BadgeHtml "HA Enabled" $EC.green) } else { 
                $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$cl.Name; Time=$nowUtcStr; Message="클러스터 HA 비활성화" }
                (Get-BadgeHtml "HA Disabled" $EC.red) 
            }
            $drsStatus = if ($cl.DrsEnabled) { (Get-BadgeHtml "DRS ($($cl.DrsAutomationLevel))" $EC.green) } else { 
                $CategorizedAlerts["Resource"] += [PSCustomObject]@{ VC=$vc; Target=$cl.Name; Time=$nowUtcStr; Message="클러스터 DRS 비활성화" }
                (Get-BadgeHtml "DRS Disabled" $EC.yellow) 
            }
            $evcMode = if ($cl.EVCMode) { $cl.EVCMode } else { "Disabled / Mixed" }
            $AllClusterRows += ,@("<b>$vc</b>", "<b>$($cl.Name)</b>", $haStatus, $drsStatus, $evcMode)
        }

        # 1. ESXi Hosts Status
        foreach ($h in $allHosts) {
            $hMoRef = $h.MoRef.Value
            if ($h.Runtime.ConnectionState -ne "connected") { 
                $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="호스트 연결 단절 ($($h.Runtime.ConnectionState))" }
            }
            if ($h.Runtime.InMaintenanceMode) {
                $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="유지보수 모드(Maintenance Mode) 상태 방치" }
            }
            if (-not $h.Config.DateTimeInfo.NtpConfig.Server) {
                $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="NTP 시간 동기화 미설정" }
            }

            $cpuMhz = if ($h.Summary.Hardware.CpuMhz) { $h.Summary.Hardware.CpuMhz * $h.Summary.Hardware.NumCpuCores } else { 0 }
            $memMB  = if ($h.Summary.Hardware.MemorySize) { $h.Summary.Hardware.MemorySize / 1MB } else { 0 }
            $cpuPct = if ($cpuMhz -gt 0 -and $h.Summary.QuickStats.OverallCpuUsage) { [math]::Round(($h.Summary.QuickStats.OverallCpuUsage / $cpuMhz) * 100, 1) } else { 0 }
            $memPct = if ($memMB -gt 0 -and $h.Summary.QuickStats.OverallMemoryUsage) { [math]::Round(($h.Summary.QuickStats.OverallMemoryUsage / $memMB) * 100, 1) } else { 0 }
            
            $cpuBar = Get-OutlookBar -pct $cpuPct; $memBar = Get-OutlookBar -pct $memPct
            $cpuStr = if ($cpuPct -ge $CpuMemCriticalPct) { "<b style='color:$($EC.red)'>$cpuPct%</b>" } else { "$cpuPct%" }
            $memStr = if ($memPct -ge $CpuMemCriticalPct) { "<b style='color:$($EC.red)'>$memPct%</b>" } else { "$memStr%" }

            $pCPU = if ($h.Hardware.CpuInfo.NumCpuCores) { $h.Hardware.CpuInfo.NumCpuCores } else { 1 }
            $pMemGB = if ($h.Hardware.MemorySize) { [math]::Round($h.Hardware.MemorySize / 1GB, 1) } else { 1 }
            
            $vCPUAssigned = if ($hostResourceMap.ContainsKey($hMoRef)) { $hostResourceMap[$hMoRef].TotalvCPU } else { 0 }
            $vMemGBAssigned = if ($hostResourceMap.ContainsKey($hMoRef)) { [math]::Round($hostResourceMap[$hMoRef].TotalvMemMB / 1024, 1) } else { 0 }
            $vmCount = if ($hostResourceMap.ContainsKey($hMoRef)) { $hostResourceMap[$hMoRef].VMCount } else { 0 }
            $onVmCount = if ($hostResourceMap.ContainsKey($hMoRef)) { $hostResourceMap[$hMoRef].OnVMCount } else { 0 }

            $cpuVirtPct = [math]::Round(($vCPUAssigned / $pCPU) * 100, 1)
            $cpuRatioStr = "<b>${cpuVirtPct}%</b> ($vCPUAssigned vCPU / $pCPU pCore)"
            if ($cpuVirtPct -ge $CpuVirtualizationPct) {
                $cpuRatioStr = "<b style='color:$($EC.red)'>$cpuRatioStr</b>"
                $CategorizedAlerts["Resource"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="CPU 가상화율 과할당 ($cpuVirtPct%)" }
            }

            $memVirtPct = [math]::Round(($vMemGBAssigned / $pMemGB) * 100, 1)
            $memOversubStr = "<b>${memVirtPct}%</b> (${vMemGBAssigned}GB / ${pMemGB}GB)"
            if ($memVirtPct -ge $MemVirtualizationPct) {
                $memOversubStr = "<b style='color:$($EC.red)'>$memOversubStr</b>"
                $CategorizedAlerts["Resource"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="메모리 가상화율 과할당 ($memVirtPct%)" }
            }

            $hwErrs = $h.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object { $_.HealthState.Key -match "red|yellow" }
            if ($hwErrs) { 
                $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="하드웨어(CIM) 센서 결함 감지" }
            }
            $stateBadge = if ($h.Runtime.ConnectionState -eq "connected") { (Get-BadgeHtml "Connected" $EC.green) } else { (Get-BadgeHtml "Disconnected" $EC.red) }

            $downNicsCount = 0
            if ($h.Config.Network.Pnic) { $downNicsCount = @($h.Config.Network.Pnic | Where-Object { $_.LinkSpeed -eq $null }).Count }
            if ($downNicsCount -gt 0) { 
                $CategorizedAlerts["Network"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="물리 NIC Down ($downNicsCount 개 Port)" }
            }
            $nicBadge = if ($downNicsCount -gt 0) { (Get-BadgeHtml "$downNicsCount NIC Down" $EC.red) } else { (Get-BadgeHtml "NICs OK" $EC.green) }

            $deadPathsCount = 0
            if ($h.Config.StorageDevice.MultipathInfo.Lun) {
                foreach ($lun in $h.Config.StorageDevice.MultipathInfo.Lun) { $deadPathsCount += @($lun.Path | Where-Object { $_.State -match "dead" }).Count }
            }
            if ($deadPathsCount -gt 0) { 
                $CategorizedAlerts["Network"] += [PSCustomObject]@{ VC=$vc; Target=$h.Name; Time=$nowUtcStr; Message="Storage Dead Path 감지 ($deadPathsCount 개 Path)" }
            }
            $deadPathBadge = if ($deadPathsCount -gt 0) { (Get-BadgeHtml "$deadPathsCount Paths Dead" $EC.red) } else { (Get-BadgeHtml "Paths OK" $EC.green) }

            $hAlarms = $h.TriggeredAlarmState | ForEach-Object { Get-AlarmName $_.Alarm $conn }
            $hAlarmStr = if ($hAlarms) { (Get-BadgeHtml "Alarm" $EC.red) + "<br><span style='color:$($EC.red);font-size:11px;'>" + ($hAlarms -join "<br>") + "</span>" } else { (Get-BadgeHtml "OK" $EC.green) }

            $AllHostRows += ,@("<b>$vc</b>", "<b>$($h.Name)</b>", $stateBadge, "<b>$vmCount 대</b><br>($onVmCount On)", $cpuRatioStr, $memOversubStr, "CPU: $cpuStr$cpuBar Mem: $memStr$memBar", "$nicBadge<br>$deadPathBadge", $hAlarmStr)
        }

        # 2. Datastores
        $dsStats = @{}
        if ($null -ne $allDatastores -and $allDatastores.Count -gt 0) {
            $stats = Get-Stat -Entity $allDatastores -Stat "datastore.totalReadLatency.average","datastore.totalWriteLatency.average" -Realtime -MaxSamples 1 -ErrorAction SilentlyContinue
            if ($stats) { $dsStats = $stats | Group-Object -Property @{Expression={$_.Entity.Name}} -AsHashTable -AsString }
        }

        foreach ($ds in $allDatastores) {
            $capGB  = [math]::Round($ds.CapacityGB, 1)
            $freeGB = [math]::Round($ds.FreeSpaceGB, 1)
            $usedGB = [math]::Round($capGB - $freeGB, 1)
            
            $freePct = if ($capGB -gt 0) { [math]::Round(($freeGB / $capGB) * 100, 1) } else { 0 }
            $usedPct = if ($capGB -gt 0) { [math]::Round(($usedGB / $capGB) * 100, 1) } else { 0 }
            
            $provGB = [math]::Round(($ds.ExtensionData.Summary.Capacity - $ds.ExtensionData.Summary.FreeSpace + $ds.ExtensionData.Summary.Uncommitted) / 1GB, 1)
            $provPct = if ($capGB -gt 0) { [math]::Round(($provGB / $capGB) * 100, 1) } else { 0 }

            if ($freePct -le 15) { 
                $CategorizedAlerts["Storage"] += [PSCustomObject]@{ VC=$vc; Target=$ds.Name; Time=$nowUtcStr; Message="잔여 용량 부족 ($freePct% 남음)" }
            }
            if ($provPct -ge 150) { 
                $CategorizedAlerts["Storage"] += [PSCustomObject]@{ VC=$vc; Target=$ds.Name; Time=$nowUtcStr; Message="Thin Provisioning Over-commit 위험 ($provPct%)" }
            }

            $rLat = 0; $wLat = 0
            if ($dsStats.ContainsKey($ds.Name)) {
                $rLat = [math]::Round(($dsStats[$ds.Name] | Where-Object MetricId -eq "datastore.totalReadLatency.average").Value, 1)
                $wLat = [math]::Round(($dsStats[$ds.Name] | Where-Object MetricId -eq "datastore.totalWriteLatency.average").Value, 1)
            }
            if ($rLat -ge $LatencyWarningMs -or $wLat -ge $LatencyWarningMs) { 
                $CategorizedAlerts["Storage"] += [PSCustomObject]@{ VC=$vc; Target=$ds.Name; Time=$nowUtcStr; Message="I/O Latency 지연 (R:${rLat}ms / W:${wLat}ms)" }
            }
            
            $storageBarHtml = Get-StorageBar -usedPct $usedPct
            $freeHtml = if ($freePct -le 15) { 
                "<b style='color:$($EC.red);'>$usedGB / $capGB GB (여유 $freePct%)</b>" 
            } else { 
                "$usedGB / $capGB GB (여유 $freePct%)" 
            }
            
            $provHtml = if ($provPct -ge 150) { (Get-BadgeHtml "$provPct%" $EC.red) } else { "$provPct%" }
            $rStr = if ($rLat -ge $LatencyWarningMs) { "<b style='color:$($EC.red)'>$rLat ms</b>" } else { "$rLat ms" }
            $wStr = if ($wLat -ge $LatencyWarningMs) { "<b style='color:$($EC.red)'>$wLat ms</b>" } else { "$wLat ms" }

            $dsAlarms = $ds.ExtensionData.TriggeredAlarmState | ForEach-Object { Get-AlarmName $_.Alarm $conn }
            $dsAlarmStr = if ($dsAlarms) { (Get-BadgeHtml "Alarm" $EC.red) + "<br><span style='color:$($EC.red);font-size:11px;'>" + ($dsAlarms -join "<br>") + "</span>" } else { (Get-BadgeHtml "OK" $EC.green) }

            $AllDatastoreRows += ,@("<b>$vc</b>", "<b>$($ds.Name)</b>", "$capGB GB", "$freeHtml $storageBarHtml", "$provHtml", "$rStr / $wStr", $dsAlarmStr)
        }

        # 3. Full VM Inventory
        foreach ($vm in ($allVMs | Sort-Object Name)) {
            $hostName = if ($vm.Runtime.Host) { $hostMap[$vm.Runtime.Host.Value] } else { "Unknown" }
            $vcpu = if ($vm.Summary.Config.NumCpu) { $vm.Summary.Config.NumCpu } else { 0 }
            $vmemGB = if ($vm.Summary.Config.MemorySizeMB) { [math]::Round($vm.Summary.Config.MemorySizeMB / 1024, 1) } else { 0 }
            $usedGB = if ($vm.Summary.Storage.Committed) { [math]::Round($vm.Summary.Storage.Committed / 1GB, 1) } else { 0 }
            $totalGB = if ($vm.Summary.Storage.Committed) { [math]::Round(($vm.Summary.Storage.Committed + $vm.Summary.Storage.Uncommitted) / 1GB, 1) } else { 0 }

            $ipList = @()
            if ($vm.Guest.Net) {
                foreach ($nic in $vm.Guest.Net) {
                    if ($nic.IpAddress) { $ipList += $nic.IpAddress | Where-Object { $_ -match "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$" } }
                }
            } elseif ($vm.Guest.IpAddress -match "^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$") { $ipList += $vm.Guest.IpAddress }
            $ipList = $ipList | Select-Object -Unique
            $ipAddrStr = if ($ipList.Count -gt 0) { $ipList -join "<br>" } else { "-" }

            $snapInfoStr = (Get-BadgeHtml "None" $EC.gray)
            if ($null -ne $vm.Snapshot) {
                $vSnaps = Get-FastSnapshotTree -tree $vm.Snapshot.RootSnapshotList -vmName $vm.Name
                $snapDetails = @()
                foreach ($s in $vSnaps) {
                    $ageDays = ((Get-Date) - $s.Created).Days
                    $sCreatedStr = $s.Created.ToUniversalTime().ToString('yyyy-MM-dd HH:mm') + " UTC"
                    $sBadge = if ($ageDays -ge $SnapAgeLimitDays) {
                        $CategorizedAlerts["Snapshot"] += [PSCustomObject]@{ VC=$vc; Target=$vm.Name; Time=$sCreatedStr; Message="장기 방치 스냅샷 '$($s.Name)' ($ageDays 일 경과)" }
                        (Get-BadgeHtml "Old ($ageDays Days)" $EC.red)
                    } else { (Get-BadgeHtml "Active ($ageDays Days)" $EC.yellow) }
                    $snapDetails += "$sBadge $($s.Name)"
                }
                $snapInfoStr = $snapDetails -join "<br>"
            }

            $vnumaBadge = (Get-BadgeHtml "vNUMA OK" $EC.gray)
            if ($vcpu -gt 8 -and $vm.Config.Hardware.NumCoresPerSocket -eq 1) { 
                $vnumaBadge = (Get-BadgeHtml "Cross-NUMA" $EC.yellow)
                $CategorizedAlerts["Resource"] += [PSCustomObject]@{ VC=$vc; Target=$vm.Name; Time=$nowUtcStr; Message="vNUMA 구성 비효율 (vCPU: $vcpu, CoresPerSocket: 1)" }
            }

            # CD-ROM / ISO 매핑 방치 VM 탐지
            if (-not $vm.Config.Template -and $vm.Config.Hardware.Device) {
                $cdDev = $vm.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualCdrom] }
                foreach ($cd in $cdDev) {
                    if ($cd.Backing -is [VMware.Vim.VirtualCdromIsoBackingInfo]) {
                        $isoPath = $cd.Backing.FileName
                        $AllIsoMountedRows += ,@("<b>$vc</b>", "<b>$($vm.Name)</b>", $isoPath)
                        
                        if ($PSBoundParameters.ContainsKey('AutoDismountISO') -and $AutoDismountISO) {
                            try {
                                Get-VM -Name $vm.Name -Server $conn | Get-CDDrive | Where-Object { $_.IsoPath -ne $null } | Set-CDDrive -NoMedia -Confirm:$false -ErrorAction Stop
                                Write-DetailedLog "[$vc] VM $($vm.Name) ISO 자동 해제 완료 (-AutoDismountISO)" "SUCCESS"
                            } catch {
                                Write-DetailedLog "[$vc] VM $($vm.Name) ISO 자동 해제 실패: $($_.Exception.Message)" "ERROR"
                            }
                        } else {
                            $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$vm.Name; Time=$nowUtcStr; Message="ISO 파일 마운트 방치 ($isoPath)" }
                        }
                    }
                }
            }

            if ($vm.Config.Template) {
                $typeBadge = (Get-BadgeHtml "Template" $EC.purple)
                $AllVmRows += ,@("<b>$vc</b>", "<b>$($vm.Name)</b>", "$typeBadge", $hostName, "${vcpu}vCPU / ${vmemGB}GB", "$usedGB GB / $totalGB GB", $ipAddrStr, "-", "-", "-", "-")
            } else {
                $stateBadge = if ($vm.Runtime.PowerState -eq "poweredOn") { (Get-BadgeHtml "VM (On)" $EC.green) } else { (Get-BadgeHtml "VM (Off)" $EC.muted) }
                $cpuPct = 0; $memPct = 0
                if ($vm.Runtime.PowerState -eq "poweredOn" -and $vm.Runtime.MaxCpuUsage -gt 0 -and $vm.Summary.QuickStats.OverallCpuUsage) {
                    $cpuPct = [math]::Round(($vm.Summary.QuickStats.OverallCpuUsage / $vm.Runtime.MaxCpuUsage) * 100, 1)
                }
                if ($vm.Runtime.PowerState -eq "poweredOn" -and $vm.Summary.Config.MemorySizeMB -gt 0 -and $vm.Summary.QuickStats.HostMemoryUsage) {
                    $memPct = [math]::Round(($vm.Summary.QuickStats.HostMemoryUsage / $vm.Summary.Config.MemorySizeMB) * 100, 1)
                }
                
                $cpuBar = Get-OutlookBar -pct $cpuPct
                $cpuStr = if ($cpuPct -ge $CpuMemCriticalPct) { "<b style='color:$($EC.red)'>$cpuPct%</b>" } elseif ($cpuPct -ge $CpuMemWarningPct) { "<b style='color:$($EC.yellow)'>$cpuPct%</b>" } else { "$cpuPct%" }
                $memStr = if ($memPct -ge $CpuMemCriticalPct) { "<b style='color:$($EC.red)'>$memStr%</b>" } else { "$memStr%" }
                if ($vm.Runtime.PowerState -ne "poweredOn") { $cpuStr = "-"; $memStr = "-"; $cpuBar = "" }

                $toolsBadge = if ($vm.Runtime.PowerState -eq "poweredOn") {
                    if ($vm.Guest.ToolsRunningStatus -eq "guestToolsNotRunning") { (Get-BadgeHtml "Not Running" $EC.red) } 
                    elseif ($vm.Guest.ToolsVersionStatus2 -match "outOfDate|needUpgrade") { (Get-BadgeHtml "Needs Upgrade" $EC.yellow) } 
                    else { (Get-BadgeHtml "Running" $EC.green) }
                } else { "-" }

                $AllVmRows += ,@("<b>$vc</b>", "<b>$($vm.Name)</b>", $stateBadge, $hostName, "${vcpu}vCPU / ${vmemGB}GB", "$usedGB GB / $totalGB GB", $ipAddrStr, "CPU: $cpuStr $cpuBar<br>Mem: $memStr", $vnumaBadge, $toolsBadge, $snapInfoStr)
            }
        }

        # 4. Administrator Security Audit
        $adminEvents = Get-VIEvent -Server $conn -Start $StartTime -MaxSamples 10000 | Where-Object { 
            $_.UserName -match "Administrator" -or $_.FullFormattedMessage -match "administrator" 
        }

        $loginEvents = $adminEvents | Where-Object { $_ -is [VMware.Vim.UserLoginSessionEvent] -or $_.FullFormattedMessage -match "logged in" -or $_.FullFormattedMessage -match "Successful login" }
        if ($loginEvents) {
            $hourlyStats = $loginEvents | Group-Object { $_.CreatedTime.ToUniversalTime().ToString("yyyy-MM-dd HH:00") } | Sort-Object Name -Descending
            foreach ($h in $hourlyStats) { $AllHourlyLoginRows += ,@("<b>$vc</b>", "<b>$($h.Name) UTC</b>", "$($h.Count) 회") }

            $sourceStats = $loginEvents | ForEach-Object {
                $ip = "Unknown IP"; $client = "Unknown Client"
                if ($_.FullFormattedMessage -match "from\s+([\d\.]+)") { $ip = $matches[1] } elseif ($_.FullFormattedMessage -match "@([\d\.]+)") { $ip = $matches[1] }
                if ($_.FullFormattedMessage -match "as\s+([^\s]+)") { $client = $matches[1] } elseif ($_.FullFormattedMessage -match "in\s+([^\s]+)") { $client = $matches[1] }
                [PSCustomObject]@{ IP = $ip; Client = $client }
            } | Group-Object IP, Client | ForEach-Object {
                $groupKey = $_.Name -split ", "
                [PSCustomObject]@{ "ClientIP" = Get-DNSHostname -ip $groupKey[0]; "ClientApp" = $groupKey[1]; "TotalCount" = $_.Count }
            } | Sort-Object "TotalCount" -Descending

            foreach ($s in $sourceStats) { $AllSourceAppRows += ,@("<b>$vc</b>", "<b>$($s.ClientIP)</b>", $s.ClientApp, "$($s.TotalCount) 회") }
        }

        $authFailEvents = $adminEvents | Where-Object { $_.FullFormattedMessage -match "failed" -or $_.FullFormattedMessage -match "Cannot login" -or $_.GetType().Name -match "BadUsername" }
        if ($authFailEvents) {
            $authFailStats = $authFailEvents | Group-Object FullFormattedMessage | ForEach-Object {
                $item = $_.Group[0]
                $failTimeStr = $item.CreatedTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + " UTC"
                [PSCustomObject]@{ "LatestTime" = $failTimeStr; "Count" = $_.Count; "Details" = $item.FullFormattedMessage }
            } | Sort-Object "LatestTime" -Descending
            
            foreach ($af in $authFailStats) {
                $CategorizedAlerts["Security"] += [PSCustomObject]@{ VC=$vc; Target="Guest OS / Auth"; Time=$af.LatestTime; Message="인증 실패 감지 ($($af.Count) 회) - $($af.Details)" }
                $AllAuthFailRows += ,@("<b>$vc</b>", $af.LatestTime, "<b style='color:$($EC.red);'>$($af.Count) 회</b>", $af.Details)
            }
        }

        $actionEvents = $adminEvents | Where-Object { $_.GetType().Name -notmatch "SessionEvent" -and $_.FullFormattedMessage -notmatch "Successful login" -and $_.FullFormattedMessage -notmatch "failed" }
        if ($actionEvents) {
            $actionStats = $actionEvents | Group-Object FullFormattedMessage | ForEach-Object {
                $item = $_.Group[0]
                [PSCustomObject]@{ "LatestTime" = $item.CreatedTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + " UTC"; "EventType" = $item.GetType().Name.Replace("Event",""); "Count" = $_.Count; "Details" = $item.FullFormattedMessage }
            } | Sort-Object "LatestTime" -Descending | Select-Object -First 15

            foreach ($act in $actionStats) { $AllAdminActionRows += ,@("<b>$vc</b>", $act.LatestTime, $act.EventType, "$($act.Count) 회", $act.Details) }
        }

        # 5. Operations Alerts
        $issueVMs = $allVMs | Where-Object { $_.Runtime.ConnectionState -match "orphaned|inaccessible|invalid" -or $_.Runtime.ConsolidationNeeded -eq $true }
        foreach ($ivm in $issueVMs) { 
            if ($ivm.Runtime.ConnectionState -match "orphaned|inaccessible|invalid") { 
                $CategorizedAlerts["Hardware"] += [PSCustomObject]@{ VC=$vc; Target=$ivm.Name; Time=$nowUtcStr; Message="VM 고아/접속불가 상태 ($($ivm.Runtime.ConnectionState))" }
            }
            if ($ivm.Runtime.ConsolidationNeeded -eq $true) { 
                $CategorizedAlerts["Storage"] += [PSCustomObject]@{ VC=$vc; Target=$ivm.Name; Time=$nowUtcStr; Message="디스크 Consolidation 필요" }
            }
            
            $stateBadge = if ($ivm.Runtime.ConnectionState -ne "connected") { (Get-BadgeHtml $ivm.Runtime.ConnectionState $EC.red) } else { "" }
            $consBadge  = if ($ivm.Runtime.ConsolidationNeeded -eq $true) { (Get-BadgeHtml "Needs Consolidation" $EC.yellow) } else { "" }
            $AllOrphanedVmRows += ,@("<b>$vc</b>", "<b style='color:$($EC.red)'>$($ivm.Name)</b>", "$stateBadge $consBadge")
        }

        $alertEvents = Get-VIEvent -Server $conn -Start $StartTime -MaxSamples 2000 | Where-Object { 
            $_.FullFormattedMessage -match "Error|Fail|HA|Warning|Alarm|down|die" -or $_ -is [VMware.Vim.VmFailedToPowerOnEvent] -or $_ -is [VMware.Vim.EventEx]
        } | Sort-Object CreatedTime -Descending | Select-Object -First 30

        foreach ($e in $alertEvents) {
            $timeStr = $e.CreatedTime.ToUniversalTime().ToString('MM-dd HH:mm') + " UTC"
            $targetStr = if ($e.Vm) { $e.Vm.Name } elseif ($e.Host) { $e.Host.Name } elseif ($e.ObjectName) { $e.ObjectName } else { "System" }
            $typeStr = $e.GetType().Name.Replace("Event","")
            $badge = if ($typeStr -match "Error|Fail") { (Get-BadgeHtml "Error" $EC.red) } elseif ($typeStr -match "Warning") { (Get-BadgeHtml "Warning" $EC.yellow) } elseif ($e.FullFormattedMessage -match "vSphere HA|High Availability") { (Get-BadgeHtml "HA" $EC.pink) } else { (Get-BadgeHtml "Alert" $EC.muted) }
            $AllCriticalEventRows += ,@("<b>$vc</b>", $timeStr, $badge, "<b>$targetStr</b>", $e.FullFormattedMessage)
        }

    } catch {
        Write-DetailedLog "$vc 수집 실패: $($_.Exception.Message)" "ERROR"
    } finally {
        if ($conn) { Disconnect-VIServer -Server $conn -Confirm:$false -WarningAction SilentlyContinue | Out-Null }
    }
}

try { $CurrentVMHistory | ConvertTo-Json -Depth 3 | Out-File $HistoryFilePath -Encoding UTF8 } catch {}

# ------------------------------------------------------------
# 5. Factor 중심 리포트 HTML 대시보드 조립
# ------------------------------------------------------------
$UtcHeaderTime = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

# 📢 Executive Summary (Attention)
$HtmlAttention = "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='margin-bottom:10px;'><tr><td style='font-size:16px;font-weight:bold;color:$($EC.red);font-family:Arial,sans-serif;border-bottom:2px solid #fca5a5;padding-bottom:5px;'>📢 전체 인프라 주요 주의사항 (Attention Executive Summary)</td></tr></table>"

$totalAlertCount = 0
foreach ($cat in $CategorizedAlerts.Keys) { $totalAlertCount += $CategorizedAlerts[$cat].Count }

if ($totalAlertCount -gt 0) {
    if ($CategorizedAlerts["Hardware"].Count -gt 0) {
        $hwRows = $CategorizedAlerts["Hardware"] | ForEach-Object { ,@($_.VC, "<b>$($_.Target)</b>", $_.Time, $_.Message) }
        $HtmlAttention += "<div style='margin:10px 0 4px 0;font-size:13px;font-weight:bold;color:$($EC.red);'>🚨 1. vCenter Appliance & Hardware Faults</div>"
        $HtmlAttention += Build-EmailTable -Headers @("vCenter", "Target Object", "Time (UTC)", "Issue / Event Details") -Rows $hwRows
    }
    if ($CategorizedAlerts["Network"].Count -gt 0) {
        $netRows = $CategorizedAlerts["Network"] | ForEach-Object { ,@($_.VC, "<b>$($_.Target)</b>", $_.Time, $_.Message) }
        $HtmlAttention += "<div style='margin:10px 0 4px 0;font-size:13px;font-weight:bold;color:$($EC.yellow);'>🔌 2. Physical Network & Storage Path Link Down</div>"
        $HtmlAttention += Build-EmailTable -Headers @("vCenter", "Host Name", "Time (UTC)", "Link / Path Status") -Rows $netRows
    }
    if ($CategorizedAlerts["Resource"].Count -gt 0) {
        $resRows = $CategorizedAlerts["Resource"] | ForEach-Object { ,@($_.VC, "<b>$($_.Target)</b>", $_.Time, $_.Message) }
        $HtmlAttention += "<div style='margin:10px 0 4px 0;font-size:13px;font-weight:bold;color:$($EC.purple);'>⚡ 3. Host Virtualization Density & vNUMA Oversubscription</div>"
        $HtmlAttention += Build-EmailTable -Headers @("vCenter", "Target Object", "Time (UTC)", "Overcommitment / Optimization Details") -Rows $resRows
    }
    if ($CategorizedAlerts["Storage"].Count -gt 0) {
        $dsAlertRows = $CategorizedAlerts["Storage"] | ForEach-Object { ,@($_.VC, "<b>$($_.Target)</b>", $_.Time, $_.Message) }
        $HtmlAttention += "<div style='margin:10px 0 4px 0;font-size:13px;font-weight:bold;color:$($EC.blue);'>💾 4. Datastore Free Space & Thin Provisioning Over-commit</div>"
        $HtmlAttention += Build-EmailTable -Headers @("vCenter", "Datastore / Target", "Time (UTC)", "Capacity / Risk Details") -Rows $dsAlertRows
    }
    if ($CategorizedAlerts["Security"].Count -gt 0) {
        $secRows = $CategorizedAlerts["Security"] | ForEach-Object { ,@($_.VC, "<b>$($_.Target)</b>", $_.Time, $_.Message) }
        $HtmlAttention += "<div style='margin:10px 0 4px 0;font-size:13px;font-weight:bold;color:$($EC.pink);'>🔐 5. Security & Guest Operation Authentication Failures</div>"
        $HtmlAttention += Build-EmailTable -Headers @("vCenter", "Target Domain", "Latest Time (UTC)", "Authentication Failure Log") -Rows $secRows
    }
    if ($CategorizedAlerts["Snapshot"].Count -gt 0) {
        $snapRows = $CategorizedAlerts["Snapshot"] | ForEach-Object { ,@($_.VC, "<b>$($_.Target)</b>", $_.Time, $_.Message) }
        $HtmlAttention += "<div style='margin:10px 0 4px 0;font-size:13px;font-weight:bold;color:$($EC.muted);'>📸 6. Unattended Long-term Snapshots (> 7 Days)</div>"
        $HtmlAttention += Build-EmailTable -Headers @("vCenter", "VM Name", "Created Time (UTC)", "Snapshot Details") -Rows $snapRows
    }
} else {
    $HtmlAttention += "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='background-color:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;margin-bottom:30px;'><tr><td style='padding:15px;color:#166534;font-size:13px;font-weight:bold;font-family:Arial,sans-serif;'>"
    $HtmlAttention += "✅ 전체 데이터센터 인프라에 장애 징후 및 집적도 과할당이 없습니다."
    $HtmlAttention += "</td></tr></table>"
}

# ------------------------------------------------------------
# 💡 [검증 팩터(Factor) 중심 통합 본문 레이아웃]
# ------------------------------------------------------------
$HtmlDetails = ""

# 📊 1. VM Inventory Trends
$HtmlDetails += Build-SectionHeader "📊" "1. 전일 대비 VM 증감 추적 (VM Inventory Trends)"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "전일 VM 수", "금일 VM 수", "증감 (Diff)", "신규 생성 VM (+)", "삭제/미검출 VM (-)") -Rows $AllVmTrendRows

# 🏢 0. vCenter Appliance Status & Partitions & Cluster
$HtmlDetails += Build-SectionHeader "🏢" "0. vCenter Appliance Status & Partitions"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "인증서 만료기한", "vCenter 트리거 알람") -Rows $AllVcStatusRows
$HtmlDetails += "<div style='margin:10px 0 6px 0;font-weight:bold;color:$($EC.text);font-size:13px;'>▶ VCSA Appliance Disk Partitions (/storage/log, /storage/db 등)</div>"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "마운트 지점", "전체 용량", "사용량", "여유 용량", "사용률 (%)", "파일시스템") -Rows $AllVcPartitionRows -EmptyMsg "vCenter 파티션 정보를 로드할 수 없습니다."
$HtmlDetails += "<div style='margin:15px 0 6px 0;font-weight:bold;color:$($EC.text);font-size:13px;'>▶ Cluster HA / DRS / EVC Mode Readiness (최근 24시간 vMotion 발생: $vMotionTotalCount 건)</div>"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Cluster Name", "HA Status", "DRS Status", "EVC Mode") -Rows $AllClusterRows -EmptyMsg "클러스터 정보가 없습니다."

# 🖥️ 1. ESXi Hosts Status
$HtmlDetails += Build-SectionHeader "🖥️" "1. ESXi Hosts Status & Virtualization Density (%)"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Host Name", "State", "VM Density", "CPU 가상화율 (%)", "Mem 가상화율 (%)", "Realtime Usage", "NIC / Paths", "Host Alarms") -Rows $AllHostRows

# 💾 2. Datastores & Thin Provisioning Risk
$HtmlDetails += Build-SectionHeader "💾" "2. Datastores & Thin Provisioning Risk"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Datastore", "Capacity", "Usage / Free Space", "Over-Prov.", "Latency (Read/Write)", "DS Alarms") -Rows $AllDatastoreRows

# 🧱 3. Full VM Inventory & Snapshot Status
$HtmlDetails += Build-SectionHeader "🧱" "3. Full VM Inventory & Snapshot Status"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "VM Name", "Type / State", "Host Name", "Allocated (CPU/Mem)", "Storage (Used/Total)", "IP Address List", "CPU / Mem Usage(%)", "vNUMA", "VMware Tools", "Snapshot Status") -Rows $AllVmRows
if ($AllIsoMountedRows) {
    $HtmlDetails += "<div style='margin:15px 0 6px 0;font-weight:bold;color:$($EC.text);font-size:13px;'>▶ Attached ISO / CD-ROM Mounted VMs (vMotion Fail Driver)</div>"
    $HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "VM Name", "Mounted ISO File Path") -Rows $AllIsoMountedRows
}

# 🔐 4. Administrator Security & Activity Audit
$HtmlDetails += Build-SectionHeader "🔐" "4. Administrator Security & Activity Audit (24H)"
$HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>▶ [4-1] Hourly Administrator Login Trends</div>"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Time Bucket (Hour)", "Login Count") -Rows $AllHourlyLoginRows -EmptyMsg "최근 24시간 내 로그인 이력이 없습니다."
$HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>▶ [4-2] Client IP / Hostname & Application Summary</div>"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Client IP & Hostname", "Client App / Protocol", "Total Login Count") -Rows $AllSourceAppRows -EmptyMsg "접속 출처 데이터가 없습니다."
$HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>▶ [4-3] Security Alerts & Authentication Failures</div>"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Latest Time (UTC)", "Failure Count", "Failure Details") -Rows $AllAuthFailRows -EmptyMsg "✔️ 최근 24시간 내 감지된 인증 실패 이력이 없습니다."
$HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>▶ [4-4] Executed Administrative Management Tasks</div>"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Time (UTC)", "Event Type", "Count", "Task Message / Details") -Rows $AllAdminActionRows -EmptyMsg "최근 수집된 순수 관리 작업 이벤트가 없습니다."

# 🚨 5. Operations Alerts
$HtmlDetails += Build-SectionHeader "🚨" "5. Operations Alerts"
if ($AllOrphanedVmRows) {
    $HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>📌 Orphaned, Inaccessible or Needs Consolidation VMs</div>"
    $HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "VM Name", "Issue Detected") -Rows $AllOrphanedVmRows
}
$HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>📌 Recent Critical Events & Alerts (Top 30)</div>"
$HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "Time (UTC)", "Type", "Target", "Message / Details") -Rows $AllCriticalEventRows -EmptyMsg "최근 24시간 내 감지된 경고 및 장애 이벤트가 없습니다. ✔️"

# ------------------------------------------------------------
# 6. HTML 최종 조합 및 저장
# ------------------------------------------------------------
$JsFilterScript = @"
<script>
document.addEventListener("DOMContentLoaded", function() {
    var searchInput = document.getElementById("reportSearchInput");
    if(searchInput) {
        searchInput.addEventListener("keyup", function() {
            var filter = this.value.toLowerCase();
            var rows = document.querySelectorAll(".data-row");
            rows.forEach(function(row) {
                var text = row.innerText.toLowerCase();
                row.style.display = text.indexOf(filter) > -1 ? "" : "none";
            });
        });
    }
});
</script>
"@

$HtmlHeader = @"
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:$($EC.bg);padding:20px;font-family:Arial,sans-serif;">
<tr><td align="center">
<table width="1000" cellpadding="0" cellspacing="0" border="0" style="background-color:$($EC.surface);border:1px solid $($EC.border);border-radius:8px;width:1000px;">
<tr><td style="background-color:$($EC.th_bg);padding:25px;border-radius:8px 8px 0 0;">
    <h1 style="color:#ffffff;margin:0;font-size:22px;">🛡️ VMware Infrastructure Master O&M Health Check</h1>
    <p style="color:#94a3b8;margin:5px 0 0 0;font-size:12px;">Standard Generation Time: $UtcHeaderTime UTC | Target Period: Last 24 Hours</p>
    <div style="margin-top:15px;">
        <input type="text" id="reportSearchInput" placeholder="🔍 Search Host, VM Name, IP or Event..." style="padding:8px 12px;width:300px;border-radius:4px;border:none;font-size:12px;">
    </div>
</td></tr>
<tr><td style="padding:25px;">
"@

$HtmlFooter = "</td></tr></table></td></tr></table>$JsFilterScript" 
$FinalHtml = "<!DOCTYPE html><html lang='ko'><head><meta charset='UTF-8'></head><body>" + $HtmlHeader + $HtmlAttention + $HtmlDetails + $HtmlFooter + "</body></html>"

# 💡 동적 파일명 생성 (DailyHealthCheck_YYYYMMDD_HHMM.html)
$TimeStamp = Get-Date -Format 'yyyyMMdd_HHmm'
$TargetFileName = "DailyHealthCheck_$TimeStamp.html"
$LocalHtmlPath = "$PSScriptRoot\$TargetFileName"

# 1. 로컬에 오늘 날짜 타임스탬프 파일 저장
$FinalHtml | Out-File -FilePath $LocalHtmlPath -Encoding UTF8 -Force

# 2. GitHub Pages용 호스팅 index.html 동시 갱신
$IndexHtmlPath = "$PSScriptRoot\index.html"
$FinalHtml | Out-File -FilePath $IndexHtmlPath -Encoding UTF8 -Force

Write-DetailedLog "로컬 HTML 리포트 저장 완료: $TargetFileName" "SUCCESS"

# ------------------------------------------------------------
# 7. SMTP 이메일 발송 (Pure PowerShell)
# ------------------------------------------------------------
Write-DetailedLog "Gmail SMTP 이메일 발송 시작..." "INFO"
$cleanPass = if ($SmtpPass) { $SmtpPass.Replace(" ", "") } else { "" }

try {
    $socket = New-Object System.Net.Sockets.TcpClient
    $connectTask = $socket.ConnectAsync($SmtpServer, $SmtpPort)
    if (-not $connectTask.Wait(3000)) {
        throw "SMTP 서버 통신 타임아웃: ${SmtpServer}:${SmtpPort} 포트에 접근할 수 없습니다."
    }
    $socket.Close()

    $mailMsg = New-Object System.Net.Mail.MailMessage
    $mailMsg.From = New-Object System.Net.Mail.MailAddress($From)
    $mailMsg.To.Add($From)
    foreach ($bcc in $BccRecipients) { if ($bcc) { $mailMsg.Bcc.Add($bcc) } }
    
    $mailMsg.Subject = $Subject
    $mailMsg.Body = $FinalHtml
    $mailMsg.IsBodyHtml = $true
    $mailMsg.BodyEncoding = [System.Text.Encoding]::UTF8
    $mailMsg.SubjectEncoding = [System.Text.Encoding]::UTF8

    $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $smtpClient.EnableSsl = $true
    $smtpClient.Timeout = 20000
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $cleanPass)

    $smtpClient.Send($mailMsg)
    $mailMsg.Dispose()

    Write-DetailedLog "이메일 발송 완료! ($($BccRecipients.Count)명 수신자 전달)" "SUCCESS"
} catch {
    Write-DetailedLog "이메일 발송 실패: $($_.Exception.Message)" "ERROR"
}

# ------------------------------------------------------------
# 8. GitHub REST API Dynamic 업로드 (v12.2 통합 적용)
# ------------------------------------------------------------
Write-DetailedLog "GitHub REST API를 통한 레포지토리 업로드를 시작합니다..." "INFO"

if ($GitHubToken -and $GitHubRepo -and $GitHubToken -notmatch "ghp_xxxx") {
    try {
        # 💡 v12.2 로직: 오늘 날짜 타임스탬프 파일명으로 GitHub에 직접 등록
        $GitHubFileName = $TargetFileName
        $ApiUrl = "https://api.github.com/repos/$GitHubRepo/contents/$GitHubFileName"
        
        # UTF-8 바이트 변환 후 Base64 인코딩
        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($FinalHtml)
        $Base64Content = [System.Convert]::ToBase64String($Bytes, [System.Base64FormattingOptions]::None)

        $Headers = @{
            "Authorization" = "token $GitHubToken"
            "Accept"        = "application/vnd.github.v3+json"
            "User-Agent"    = "PowerShell-Automation"
        }

        # 1. 동일한 파일명의 SHA 커밋 해시 수집
        $Sha = $null
        try {
            $ExistFile = Invoke-RestMethod -Uri $ApiUrl -Headers $Headers -Method Get -ErrorAction Stop
            if ($ExistFile -and $ExistFile.sha) {
                $Sha = $ExistFile.sha
            }
        } catch {}

        # 2. Payload 구성
        $BodyObject = @{
            message = "Upload $GitHubFileName ($UtcToday UTC)"
            content = $Base64Content
        }
        if ($Sha) { $BodyObject["sha"] = $Sha }

        $JsonBody = $BodyObject | ConvertTo-Json -Depth 3

        # 3. GitHub REST API PUT 요청
        $Response = Invoke-RestMethod -Uri $ApiUrl -Headers $Headers -Method Put -Body ([System.Text.Encoding]::UTF8.GetBytes($JsonBody)) -ContentType "application/json; charset=utf-8" -ErrorAction Stop

        Write-DetailedLog "GitHub 레포지토리($GitHubRepo) 업로드 성공! 파일명: $GitHubFileName" "SUCCESS"
    } catch {
        $ghErr = $_.Exception
        Write-DetailedLog "GitHub REST API 업로드 실패: $($ghErr.Message)" "ERROR"
        if ($ghErr.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($ghErr.Response.GetResponseStream())
                $responseBody = $reader.ReadToEnd()
                Write-DetailedLog "  👉 GitHub 상세 응답: $responseBody" "WARN"
            } catch {}
        }
    }
} else {
    Write-DetailedLog "GitHubToken 또는 GitHubRepo 설정이 유효하지 않아 GitHub 업로드를 건너뜁니다." "WARN"
}

Write-DetailedLog "vSphere Master O&M Health Check v24.0 종료." "INFO"
