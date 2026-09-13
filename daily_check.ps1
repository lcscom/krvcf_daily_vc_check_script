# ====================================================================
# VMware Daily Morning Health Check - Enterprise Pure Edition v12.2
# (Features: Native Event Audit, VCSA Partitions, Timestamp File Format,
#            Detailed SMTP Diagnostics, Pure PowerShell BCC & GitHub Dynamic Upload)
# ====================================================================

$ErrorActionPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

Import-Module VMware.PowerCLI -ErrorAction SilentlyContinue
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCEIP $false -Scope Session -Confirm:$false -WarningAction SilentlyContinue | Out-Null
# ------------------------------------------------------------
# 1. 환경, 계정 및 자격 증명 설정
# ------------------------------------------------------------
# vCenter 계정 정보
$vcAccounts = @{
    "vcsa01.rangers.lab"       = @{ User = "administrator@vsphere.local"; Pass = "" }
    "kr-vcs9-mgmt.rangers.lab" = @{ User = "administrator@vsphere.local"; Pass = "" }
}

# Gmail SMTP 설정 (Google 계정에서 발급받은 16자리 앱 비밀번호 사용)
$SmtpServer = "smtp.gmail.com"
$SmtpPort   = 587
$SmtpUser   = ""
$SmtpPass   = ""       # 16자리 앱 비밀번호 입력 (공백 자동 제거 처리됨)

# 이메일 수신자 리스트 (BCC 숨은 참조 발송)
$BccRecipients = @(
    "",
    ""
)

# GitHub REST API 연동 설정
$GitHubToken = "" # GitHub Personal Access Token (PAT)
$GitHubRepo  = ""       # GitHub 레포지토리 (username/repo)

# 기본 실행 변수
$From       = $SmtpUser
$UtcToday   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd')
$Subject    = "🛡️ VMware Infrastructure Daily Health Check ($UtcToday UTC)"
$StartTime  = (Get-Date).AddHours(-24).ToUniversalTime()

$HistoryFilePath = "$PSScriptRoot\vm_inventory_history.json"
$LogFilePath     = "$PSScriptRoot\HealthCheck_Execution.log"

# 디테일 로거 함수
function Write-DetailedLog {
    param([string]$Message, [string]$Level = "INFO")
    $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timeStr] [$Level] $Message"
    
    switch ($Level) {
        "ERROR" { Write-Host $logLine -ForegroundColor Red }
        "WARN"  { Write-Host $logLine -ForegroundColor Yellow }
        "SUCCESS"{ Write-Host $logLine -ForegroundColor Green }
        Default { Write-Host $logLine -ForegroundColor Gray }
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
    param([double]$pct, [switch]$ReverseColor)
    if ($pct -gt 100) { $pct = 100 }
    if ($pct -lt 0) { $pct = 0 }
    $remain = 100 - $pct
    
    if ($ReverseColor) {
        $color = if ($pct -le (100-$CpuMemCriticalPct)) { $EC.red } elseif ($pct -le (100-$CpuMemWarningPct)) { $EC.yellow } else { $EC.green }
    } else {
        $color = if ($pct -ge $CpuMemCriticalPct) { $EC.red } elseif ($pct -ge $CpuMemWarningPct) { $EC.yellow } else { $EC.green }
    }
    
    $barHtml = "<table width='100px' border='0' cellpadding='0' cellspacing='0' style='border:1px solid #dde1e8;background-color:#f1f5f9;margin-top:4px;'><tr>"
    if ($pct -gt 0) { $barHtml += "<td width='$pct%' style='background-color:$color;height:6px;font-size:1px;line-height:1px;'>&nbsp;</td>" }
    if ($remain -gt 0) { $barHtml += "<td width='$remain%' style='background-color:transparent;height:6px;font-size:1px;line-height:1px;'>&nbsp;</td>" }
    $barHtml += "</tr></table>"
    return $barHtml
}

function Build-EmailTable {
    param([string[]]$Headers, $Rows, [string]$EmptyMsg = "데이터가 없습니다.")
    $thHtml = ($Headers | ForEach-Object { "<th align='left' valign='middle' style='padding:8px 10px;background-color:$($EC.th_bg);color:$($EC.th_text);font-size:12px;font-family:Arial,sans-serif;border:1px solid $($EC.border);white-space:nowrap;'>$_</th>" }) -join ""
    
    if ($Rows.Count -eq 0) {
        $bodyHtml = "<tr><td colspan='$($Headers.Count)' align='center' valign='middle' style='padding:15px;background-color:$($EC.surface);border:1px solid $($EC.border);color:$($EC.muted);font-size:12px;font-family:Arial,sans-serif;'>$EmptyMsg</td></tr>"
    } else {
        $bodyHtml = ($Rows | ForEach-Object {
            $tds = ($_ | ForEach-Object { "<td align='left' valign='top' style='padding:8px 10px;background-color:$($EC.surface);border:1px solid $($EC.border);font-size:12px;color:$($EC.text);font-family:Arial,sans-serif;word-break:break-all;'>$_</td>" }) -join ""
            "<tr>$tds</tr>"
        }) -join ""
    }
    return "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;margin-bottom:20px;width:100%;'><thead><tr>$thHtml</tr></thead><tbody>$bodyHtml</tbody></table>"
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
Write-DetailedLog "vSphere Daily Health Check 스크립트 실행 시작" "INFO"

$GlobalAlerts = @()   
$HtmlDetails  = ""    

$PrevVMHistory = @{}
if (Test-Path $HistoryFilePath) {
    try { $PrevVMHistory = Get-Content $HistoryFilePath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
}
$CurrentVMHistory = @{}
$VmTrendRows = @()

foreach ($vc in $vcAccounts.Keys) {
    Write-DetailedLog "[$vc] 데이터 수집 및 보안 감사 분석 진행 중..." "INFO"

    $conn = $null
    try {
        $user = $vcAccounts[$vc].User
        $pass = $vcAccounts[$vc].Pass
        $conn = Connect-VIServer -Server $vc -User $user -Password $pass -WarningAction SilentlyContinue -ErrorAction Stop
        
        $HtmlDetails += "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='margin-top:30px;width:100%;'><tr><td style='background-color:#1e293b;padding:12px;font-size:16px;font-weight:bold;border-radius:5px;color:#ffffff;font-family:Arial,sans-serif;'>📌 Target vCenter: $vc</td></tr></table>"

        $global:alarmCache = @{}
        function Get-AlarmName($alarmMoRef, $connection) {
            if ($null -eq $alarmMoRef -or -not $alarmMoRef.Value) { return "Unknown Alarm" }
            if (-not $global:alarmCache.ContainsKey($alarmMoRef.Value)) {
                try {
                    $alarmView = Get-View -Id $alarmMoRef -Property Info.Name -Server $connection -ErrorAction Stop
                    if ($alarmView -and $alarmView.Info) {
                        $global:alarmCache[$alarmMoRef.Value] = $alarmView.Info.Name
                    } else {
                        $global:alarmCache[$alarmMoRef.Value] = "Alarm ($($alarmMoRef.Value))"
                    }
                } catch { $global:alarmCache[$alarmMoRef.Value] = "Alarm ($($alarmMoRef.Value))" }
            }
            return $global:alarmCache[$alarmMoRef.Value]
        }

        # --- 데이터 일괄 추출 ---
        $allVMs = Get-View -ViewType VirtualMachine -Property Name, Runtime.Host, Runtime.PowerState, Runtime.ConnectionState, Snapshot, Runtime.MaxCpuUsage, Summary.QuickStats, Runtime.ConsolidationNeeded, Guest.ToolsRunningStatus, Guest.ToolsVersionStatus2, Config.Template, Summary.Config.NumCpu, Summary.Config.MemorySizeMB, Summary.Storage.Committed, Summary.Storage.Uncommitted, Guest.IpAddress, Guest.Net, Config.Hardware.NumCoresPerSocket, Guest.Disk -Server $conn
        $allHosts = Get-View -ViewType HostSystem -Property Name, Runtime.ConnectionState, Runtime.BootTime, Runtime.HealthSystemRuntime, Summary, Config, TriggeredAlarmState, Hardware.CpuInfo.NumCpuCores, Hardware.MemorySize -Server $conn
        $allDatastores = Get-Datastore -Server $conn
        $hostMap = @{}; foreach ($h in $allHosts) { $hostMap[$h.MoRef.Value] = $h.Name }

        # --- 📊 [VM 증감 추적] ---
        $currVmNames = $allVMs | Select-Object -ExpandProperty Name
        $CurrentVMHistory[$vc] = $currVmNames
        $prevVmNames = if ($PrevVMHistory.PSObject.Properties[$vc]) { $PrevVMHistory.$vc } else { @() }
        
        $newVMs = $currVmNames | Where-Object { $_ -notin $prevVmNames }
        $deletedVMs = $prevVmNames | Where-Object { $_ -notin $currVmNames }
        
        $diffCount = $currVmNames.Count - $prevVmNames.Count
        $diffStr = if ($diffCount -gt 0) { "<span style='color:$($EC.purple);font-weight:bold;'>+$diffCount 대</span>" } elseif ($diffCount -lt 0) { "<span style='color:$($EC.red);font-weight:bold;'>$diffCount 대</span>" } else { "변동 없음 (0)" }
        
        $newStr = if ($newVMs) { ($newVMs | ForEach-Object { (Get-BadgeHtml "+" $EC.purple) + " $_" }) -join "<br>" } else { "-" }
        $delStr = if ($deletedVMs) { ($deletedVMs | ForEach-Object { (Get-BadgeHtml "-" $EC.red) + " $_" }) -join "<br>" } else { "-" }

        if ($newVMs) { $GlobalAlerts += "📈 <b>[$vc]</b> 신규 VM $($newVMs.Count) 대 생성됨 ($($newVMs -join ', '))" }
        if ($deletedVMs) { $GlobalAlerts += "📉 <b>[$vc]</b> VM $($deletedVMs.Count) 대 삭제/미검출됨 ($($deletedVMs -join ', '))" }

        $VmTrendRows += ,@("<b>$vc</b>", "$($prevVmNames.Count) 대", "$($currVmNames.Count) 대", $diffStr, $newStr, $delStr)

        # --- [0] vCenter 상태 및 파티션 용량 ---
        $HtmlDetails += Build-SectionHeader "🏢" "0. vCenter Appliance Status & Partitions"
        $vcCertDays = "N/A"
        try {
            $req = [Net.HttpWebRequest]::Create("https://$vc"); $req.Timeout = 3000; $req.Method = "HEAD"; $req.GetResponse() | Out-Null
            if ($req.ServicePoint.Certificate) { $vcCertDays = ([datetime]::Parse($req.ServicePoint.Certificate.GetExpirationDateString()) - (Get-Date)).Days }
        } catch {}
        
        if ($vcCertDays -ne "N/A" -and $vcCertDays -le $CertWarningDays) { $GlobalAlerts += "🚨 <b>[$vc]</b> vCenter 인증서 만료 임박 (D-$vcCertDays 일)" }
        $vcCertBadge = if ($vcCertDays -le $CertWarningDays -and $vcCertDays -ne "N/A") { (Get-BadgeHtml "D-$vcCertDays" $EC.red) } else { (Get-BadgeHtml "D-$vcCertDays" $EC.green) }
        
        # VCSA 디스크 파티션 매핑
        $vcsaPartitionRows = @()
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
                        $GlobalAlerts += "🚨 <b>[$vc]</b> vCenter 파티션 $pPath 용량 위험 ($usedPct% 사용 중)"
                        "<b style='color:$($EC.red);'>$usedPct% (경고)</b>"
                    } else {
                        "$usedPct%"
                    }
                    $vcsaPartitionRows += ,@("<b>$pPath</b>", "$capGB GB", "$usedGB GB", "$freeGB GB", $pctDisplay, "ext4/xfs")
                }
            }
        } catch {}

        $si = Get-View -Id "ServiceInstance" -Server $conn
        $rootFolder = Get-View -Id $si.Content.RootFolder -Property TriggeredAlarmState -Server $conn
        $vcAlarmStr = ""
        if ($rootFolder -and $rootFolder.TriggeredAlarmState) {
            $vcAlarms = $rootFolder.TriggeredAlarmState | ForEach-Object { Get-AlarmName $_.Alarm $conn }
            $vcAlarmStr = "<span style='color:$($EC.red);font-weight:bold;'>" + ($vcAlarms -join "<br>") + "</span>"
            $GlobalAlerts += "🚨 <b>[$vc]</b> vCenter 어플라이언스 알람 감지 ($($vcAlarms.Count) 건)"
        } else { $vcAlarmStr = (Get-BadgeHtml "OK" $EC.green) }
        
        $HtmlDetails += Build-EmailTable -Headers @("vCenter FQDN", "인증서 만료기한", "vCenter 트리거 알람") -Rows @(,@($vc, $vcCertBadge, $vcAlarmStr))
        
        $HtmlDetails += "<div style='margin:10px 0 6px 0;font-weight:bold;color:$($EC.text);font-family:Arial,sans-serif;font-size:13px;'>▶ VCSA Appliance Disk Partitions (/storage/log, /storage/db, / 등)</div>"
        $HtmlDetails += Build-EmailTable -Headers @("마운트 지점 (Partition)", "전체 용량", "사용량", "여유 용량", "사용률 (%)", "파일시스템") -Rows $vcsaPartitionRows -EmptyMsg "vCenter 파티션 정보를 로드할 수 없습니다."

        # --- [1] ESXi Host 상태 및 가상화율 (%) ---
        $HtmlDetails += Build-SectionHeader "🖥️" "1. ESXi Hosts Status & Virtualization Density (%)"
        $hRows = @()
        
        $hostResourceMap = @{}
        foreach ($vm in $allVMs) {
            if (-not $vm.Config.Template -and $vm.Runtime.Host) {
                $hMoRef = $vm.Runtime.Host.Value
                if (-not $hostResourceMap.ContainsKey($hMoRef)) { $hostResourceMap[$hMoRef] = @{ TotalvCPU = 0; TotalvMemMB = 0; VMCount = 0; OnVMCount = 0 } }
                $hostResourceMap[$hMoRef].TotalvCPU += if ($vm.Summary.Config.NumCpu) { $vm.Summary.Config.NumCpu } else { 0 }
                $hostResourceMap[$hMoRef].TotalvMemMB += if ($vm.Summary.Config.MemorySizeMB) { $vm.Summary.Config.MemorySizeMB } else { 0 }
                $hostResourceMap[$hMoRef].VMCount += 1
                if ($vm.Runtime.PowerState -eq "poweredOn") { $hostResourceMap[$hMoRef].OnVMCount += 1 }
            }
        }

        foreach ($h in $allHosts) {
            $hMoRef = $h.MoRef.Value
            $uptime = if ($h.Runtime.BootTime) { "$([math]::Round(((Get-Date) - $h.Runtime.BootTime).TotalDays, 1)) days" } else { "N/A" }
            $ntp = if ($h.Config.DateTimeInfo.NtpConfig.Server) { (Get-BadgeHtml "Set" $EC.green) } else { (Get-BadgeHtml "Unset" $EC.yellow) }
            if ($h.Runtime.ConnectionState -ne "connected") { $GlobalAlerts += "🚨 <b>[$vc]</b> $($h.Name) 호스트 연결 단절 ($($h.Runtime.ConnectionState))" }

            $cpuMhz = if ($h.Summary.Hardware.CpuMhz) { $h.Summary.Hardware.CpuMhz * $h.Summary.Hardware.NumCpuCores } else { 0 }
            $memMB  = if ($h.Summary.Hardware.MemorySize) { $h.Summary.Hardware.MemorySize / 1MB } else { 0 }
            $cpuPct = if ($cpuMhz -gt 0 -and $h.Summary.QuickStats.OverallCpuUsage) { [math]::Round(($h.Summary.QuickStats.OverallCpuUsage / $cpuMhz) * 100, 1) } else { 0 }
            $memPct = if ($memMB -gt 0 -and $h.Summary.QuickStats.OverallMemoryUsage) { [math]::Round(($h.Summary.QuickStats.OverallMemoryUsage / $memMB) * 100, 1) } else { 0 }
            
            $cpuBar = Get-OutlookBar -pct $cpuPct; $memBar = Get-OutlookBar -pct $memPct
            $cpuStr = if ($cpuPct -ge $CpuMemCriticalPct) { "<b style='color:$($EC.red)'>$cpuPct%</b>" } else { "$cpuPct%" }
            $memStr = if ($memPct -ge $CpuMemCriticalPct) { "<b style='color:$($EC.red)'>$memStr%</b>" } else { "$memStr%" }

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
                $GlobalAlerts += "⚠️ <b>[$vc]</b> 호스트 $($h.Name) CPU 가상화율 과할당 ($cpuVirtPct%)"
            }

            $memVirtPct = [math]::Round(($vMemGBAssigned / $pMemGB) * 100, 1)
            $memOversubStr = "<b>${memVirtPct}%</b> (${vMemGBAssigned}GB / ${pMemGB}GB)"
            if ($memVirtPct -ge $MemVirtualizationPct) {
                $memOversubStr = "<b style='color:$($EC.red)'>$memOversubStr</b>"
                $GlobalAlerts += "⚠️ <b>[$vc]</b> 호스트 $($h.Name) 메모리 가상화율 과할당 ($memVirtPct%)"
            }

            $hwErrs = $h.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo | Where-Object { $_.HealthState.Key -match "red|yellow" }
            if ($hwErrs) { $GlobalAlerts += "🚨 <b>[$vc]</b> $($h.Name) 하드웨어(CIM) 센서 결함" }
            $hwBadge = if ($hwErrs) { (Get-BadgeHtml "Fault" $EC.red) } else { (Get-BadgeHtml "OK" $EC.green) }
            $stateBadge = if ($h.Runtime.ConnectionState -eq "connected") { (Get-BadgeHtml "Connected" $EC.green) } else { (Get-BadgeHtml "Disconnected" $EC.red) }

            $downNicsCount = 0
            if ($h.Config.Network.Pnic) { $downNicsCount = @($h.Config.Network.Pnic | Where-Object { $_.LinkSpeed -eq $null }).Count }
            if ($downNicsCount -gt 0) { $GlobalAlerts += "🚨 <b>[$vc]</b> $($h.Name) vmnic $downNicsCount 개 Link Down" }
            $nicBadge = if ($downNicsCount -gt 0) { (Get-BadgeHtml "$downNicsCount NIC Down" $EC.red) } else { (Get-BadgeHtml "NICs OK" $EC.green) }

            $deadPathsCount = 0
            if ($h.Config.StorageDevice.MultipathInfo.Lun) {
                foreach ($lun in $h.Config.StorageDevice.MultipathInfo.Lun) { $deadPathsCount += @($lun.Path | Where-Object { $_.State -match "dead" }).Count }
            }
            if ($deadPathsCount -gt 0) { $GlobalAlerts += "🚨 <b>[$vc]</b> $($h.Name) Dead Paths $deadPathsCount 개 감지" }
            $deadPathBadge = if ($deadPathsCount -gt 0) { (Get-BadgeHtml "$deadPathsCount Paths Dead" $EC.red) } else { (Get-BadgeHtml "Paths OK" $EC.green) }

            $hAlarms = $h.TriggeredAlarmState | ForEach-Object { Get-AlarmName $_.Alarm $conn }
            $hAlarmStr = if ($hAlarms) { (Get-BadgeHtml "Alarm" $EC.red) + "<br><span style='color:$($EC.red);font-size:11px;'>" + ($hAlarms -join "<br>") + "</span>" } else { (Get-BadgeHtml "OK" $EC.green) }

            $hRows += ,@("<b>$($h.Name)</b>", $stateBadge, "<b>$vmCount 대</b><br>($onVmCount On)", $cpuRatioStr, $memOversubStr, "CPU: $cpuStr$cpuBar Mem: $memStr$memBar", "$nicBadge<br>$deadPathBadge", $hAlarmStr)
        }
        $HtmlDetails += Build-EmailTable -Headers @("Host Name", "State", "VM Density", "CPU 가상화율 (%)", "Mem 가상화율 (%)", "Realtime Usage", "NIC / Paths", "Host Alarms") -Rows $hRows

        # --- [2] Datastore I/O ---
        $HtmlDetails += Build-SectionHeader "💾" "2. Datastores & Latency"
        $dsRows = @()
        $dsStats = @{}
        
        if ($null -ne $allDatastores -and $allDatastores.Count -gt 0) {
            $stats = Get-Stat -Entity $allDatastores -Stat "datastore.totalReadLatency.average","datastore.totalWriteLatency.average" -Realtime -MaxSamples 1 -ErrorAction SilentlyContinue
            if ($stats) { $dsStats = $stats | Group-Object -Property @{Expression={$_.Entity.Name}} -AsHashTable -AsString }
        }

        foreach ($ds in $allDatastores) {
            $capGB  = [math]::Round($ds.CapacityGB, 1); $freeGB = [math]::Round($ds.FreeSpaceGB, 1)
            $freePct = if ($capGB -gt 0) { [math]::Round(($freeGB / $capGB) * 100, 1) } else { 0 }
            $provGB = [math]::Round(($ds.ExtensionData.Summary.Capacity - $ds.ExtensionData.Summary.FreeSpace + $ds.ExtensionData.Summary.Uncommitted) / 1GB, 1)
            $provPct = if ($capGB -gt 0) { [math]::Round(($provGB / $capGB) * 100, 1) } else { 0 }

            if ($freePct -le 15) { $GlobalAlerts += "⚠️ <b>[$vc]</b> 스토리지 $($ds.Name) 잔여 용량 부족 ($freePct%)" }
            if ($provPct -ge 150) { $GlobalAlerts += "⚠️ <b>[$vc]</b> 스토리지 $($ds.Name) Over-provisioning 심각 ($provPct%)" }

            $rLat = 0; $wLat = 0
            if ($dsStats.ContainsKey($ds.Name)) {
                $rLat = [math]::Round(($dsStats[$ds.Name] | Where-Object MetricId -eq "datastore.totalReadLatency.average").Value, 1)
                $wLat = [math]::Round(($dsStats[$ds.Name] | Where-Object MetricId -eq "datastore.totalWriteLatency.average").Value, 1)
            }
            if ($rLat -ge $LatencyWarningMs -or $wLat -ge $LatencyWarningMs) { $GlobalAlerts += "⚠️ <b>[$vc]</b> $($ds.Name) I/O Latency 지연 (R:${rLat}ms/W:${wLat}ms)" }
            
            $freeBar  = Get-OutlookBar -pct $freePct -ReverseColor
            $freeHtml = if ($freePct -le 15) { "<b style='color:$($EC.red)'>$freeGB GB ($freePct%)</b>" } else { "$freeGB GB ($freePct%)" }
            $provHtml = if ($provPct -ge 150) { (Get-BadgeHtml "$provPct%" $EC.red) } else { "$provPct%" }
            
            $rStr = if ($rLat -ge $LatencyWarningMs) { "<b style='color:$($EC.red)'>$rLat ms</b>" } else { "$rLat ms" }
            $wStr = if ($wLat -ge $LatencyWarningMs) { "<b style='color:$($EC.red)'>$wLat ms</b>" } else { "$wLat ms" }

            $dsAlarms = $ds.ExtensionData.TriggeredAlarmState | ForEach-Object { Get-AlarmName $_.Alarm $conn }
            $dsAlarmStr = if ($dsAlarms) { (Get-BadgeHtml "Alarm" $EC.red) + "<br><span style='color:$($EC.red);font-size:11px;'>" + ($dsAlarms -join "<br>") + "</span>" } else { (Get-BadgeHtml "OK" $EC.green) }

            $dsRows += ,@("<b>$($ds.Name)</b>", "$capGB GB", "$freeHtml $freeBar", "$provHtml", "$rStr / $wStr", $dsAlarmStr)
        }
        $HtmlDetails += Build-EmailTable -Headers @("Datastore", "Capacity", "Free Space", "Over-Prov.", "Latency (Read/Write)", "DS Alarms") -Rows $dsRows

        # --- [3] 전체 VM 목록 (스냅샷 현황 통합) ---
        $HtmlDetails += Build-SectionHeader "🧱" "3. Full VM Inventory & Integrated Snapshot Status"
        $vmRows = @()
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
                    $sBadge = if ($ageDays -ge $SnapAgeLimitDays) {
                        $GlobalAlerts += "⚠️ <b>[$vc]</b> VM $($vm.Name) 장기 방치 스냅샷 감지 ($ageDays 일 경과)"
                        (Get-BadgeHtml "Old ($ageDays Days)" $EC.red)
                    } else {
                        (Get-BadgeHtml "Active ($ageDays Days)" $EC.yellow)
                    }
                    $snapDetails += "$sBadge $($s.Name)"
                }
                $snapInfoStr = $snapDetails -join "<br>"
            }

            $vnumaBadge = (Get-BadgeHtml "vNUMA OK" $EC.gray)
            if ($vcpu -gt 8 -and $vm.Config.Hardware.NumCoresPerSocket -eq 1) { $vnumaBadge = (Get-BadgeHtml "Cross-NUMA" $EC.yellow) }

            if ($vm.Config.Template) {
                $typeBadge = (Get-BadgeHtml "Template" $EC.purple)
                $vmRows += ,@("<b>$($vm.Name)</b>", "$typeBadge", $hostName, "${vcpu}vCPU / ${vmemGB}GB", "$usedGB GB / $totalGB GB", $ipAddrStr, "-", "-", "-", "-")
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

                $toolsBadge = ""
                if ($vm.Runtime.PowerState -eq "poweredOn") {
                    if ($vm.Guest.ToolsRunningStatus -eq "guestToolsNotRunning") { $toolsBadge = (Get-BadgeHtml "Not Running" $EC.red) } 
                    elseif ($vm.Guest.ToolsVersionStatus2 -match "outOfDate|needUpgrade") { $toolsBadge = (Get-BadgeHtml "Needs Upgrade" $EC.yellow) } 
                    else { $toolsBadge = (Get-BadgeHtml "Running" $EC.green) }
                } else { $toolsBadge = "-" }

                $vmRows += ,@("<b>$($vm.Name)</b>", $stateBadge, $hostName, "${vcpu}vCPU / ${vmemGB}GB", "$usedGB GB / $totalGB GB", $ipAddrStr, "CPU: $cpuStr $cpuBar<br>Mem: $memStr", $vnumaBadge, $toolsBadge, $snapInfoStr)
            }
        }
        $HtmlDetails += Build-EmailTable -Headers @("VM Name", "Type / State", "Host Name", "Allocated (CPU/Mem)", "Storage (Used/Total)", "IP Address List", "CPU / Mem Usage(%)", "vNUMA", "VMware Tools", "Snapshot Status") -Rows $vmRows

        # --- [4] Administrator 통합 보안 감사 ---
        $HtmlDetails += Build-SectionHeader "🔐" "4. Administrator Security & Activity Audit (24H)"
        
        $adminEvents = Get-VIEvent -Server $conn -Start $StartTime -MaxSamples 10000 | Where-Object { 
            $_.UserName -match "Administrator" -or $_.FullFormattedMessage -match "administrator" 
        }

        # Sub 1. 시간대별 Administrator 로그인 트렌드
        $HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;color:$($EC.text);font-family:Arial,sans-serif;font-size:13px;'>▶ [4-1] Hourly Administrator Login Trends</div>"
        $loginEvents = $adminEvents | Where-Object { 
            $_ -is [VMware.Vim.UserLoginSessionEvent] -or 
            $_.FullFormattedMessage -match "logged in" -or 
            $_.FullFormattedMessage -match "Successful login"
        }
        $hourlyRows = @()
        if ($loginEvents) {
            $hourlyStats = $loginEvents | Group-Object { $_.CreatedTime.ToUniversalTime().ToString("yyyy-MM-dd HH:00") } | Sort-Object Name -Descending
            foreach ($h in $hourlyStats) {
                $hourlyRows += ,@("<b>$($h.Name) UTC</b>", "$($h.Count) 회")
            }
        }
        $HtmlDetails += Build-EmailTable -Headers @("Time Bucket (Hour)", "Login Count") -Rows $hourlyRows -EmptyMsg "최근 24시간 내 로그인 이력이 없습니다."

        # Sub 2. 접속 출처 (IP & Reverse DNS) 및 클라이언트 앱
        $HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;color:$($EC.text);font-family:Arial,sans-serif;font-size:13px;'>▶ [4-2] Client IP / Hostname & Application Summary</div>"
        $sourceRows = @()
        if ($loginEvents) {
            $sourceStats = $loginEvents | ForEach-Object {
                $ip = "Unknown IP"
                $client = "Unknown Client"

                if ($_.FullFormattedMessage -match "from\s+([\d\.]+)") { $ip = $matches[1] }
                elseif ($_.FullFormattedMessage -match "@([\d\.]+)") { $ip = $matches[1] }

                if ($_.FullFormattedMessage -match "as\s+([^\s]+)") { $client = $matches[1] }
                elseif ($_.FullFormattedMessage -match "in\s+([^\s]+)") { $client = $matches[1] }

                [PSCustomObject]@{ IP = $ip; Client = $client }
            } | Group-Object IP, Client | ForEach-Object {
                $groupKey = $_.Name -split ", "
                $ipAddress = $groupKey[0]
                $dnsInfo = Get-DNSHostname -ip $ipAddress

                [PSCustomObject]@{
                    "ClientIP" = $dnsInfo
                    "ClientApp" = $groupKey[1]
                    "TotalCount" = $_.Count
                }
            } | Sort-Object "TotalCount" -Descending

            foreach ($s in $sourceStats) {
                $sourceRows += ,@("<b>$($s.ClientIP)</b>", $s.ClientApp, "$($s.TotalCount) 회")
            }
        }
        $HtmlDetails += Build-EmailTable -Headers @("Client IP & Hostname", "Client App / Protocol", "Total Login Count") -Rows $sourceRows -EmptyMsg "접속 출처 데이터가 없습니다."

        # Sub 3. 보안 경고 및 인증 실패 이력 (Authentication Failures)
        $HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;color:$($EC.text);font-family:Arial,sans-serif;font-size:13px;'>▶ [4-3] Security Alerts & Authentication Failures</div>"
        $authFailEvents = $adminEvents | Where-Object { 
            $_.FullFormattedMessage -match "failed" -or 
            $_.FullFormattedMessage -match "Cannot login" -or
            $_.GetType().Name -match "BadUsername"
        }
        $authFailRows = @()
        if ($authFailEvents) {
            $authFailStats = $authFailEvents | Group-Object FullFormattedMessage | ForEach-Object {
                $item = $_.Group[0]
                [PSCustomObject]@{
                    "LatestTime" = $item.CreatedTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + " UTC"
                    "Count"      = $_.Count
                    "Details"    = $item.FullFormattedMessage
                }
            } | Sort-Object "LatestTime" -Descending
            
            foreach ($af in $authFailStats) {
                $GlobalAlerts += "🚨 <b>[$vc]</b> 관리자 계정 로그인 인증 실패 감지 ($($af.Count) 회) - $($af.Details)"
                $authFailRows += ,@($af.LatestTime, "<b style='color:$($EC.red);'>$($af.Count) 회</b>", $af.Details)
            }
        }
        $HtmlDetails += Build-EmailTable -Headers @("Latest Time (UTC)", "Failure Count", "Failure Details") -Rows $authFailRows -EmptyMsg "✔️ 최근 24시간 내 감지된 인증 실패 이력이 없습니다."

        # Sub 4. Administrator 주요 관리 작업 (설정 변경/알람 해제 등 순수 작업)
        $HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;color:$($EC.text);font-family:Arial,sans-serif;font-size:13px;'>▶ [4-4] Executed Administrative Management Tasks</div>"
        $actionEvents = $adminEvents | Where-Object { 
            $_.GetType().Name -notmatch "SessionEvent" -and
            $_.FullFormattedMessage -notmatch "Successful login" -and
            $_.FullFormattedMessage -notmatch "failed"
        }
        $actionRows = @()
        if ($actionEvents) {
            $actionStats = $actionEvents | Group-Object FullFormattedMessage | ForEach-Object {
                $item = $_.Group[0]
                [PSCustomObject]@{
                    "LatestTime" = $item.CreatedTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss") + " UTC"
                    "EventType"  = $item.GetType().Name.Replace("Event","")
                    "Count"      = $_.Count
                    "Details"    = $item.FullFormattedMessage
                }
            } | Sort-Object "LatestTime" -Descending | Select-Object -First 15

            foreach ($act in $actionStats) {
                $actionRows += ,@($act.LatestTime, $act.EventType, "$($act.Count) 회", $act.Details)
            }
        }
        $HtmlDetails += Build-EmailTable -Headers @("Latest Time (UTC)", "Event Type", "Count", "Task Message / Details") -Rows $actionRows -EmptyMsg "최근 수집된 순수 관리 작업 이벤트가 없습니다."

        # --- [5] Operations Alerts ---
        $HtmlDetails += Build-SectionHeader "🚨" "5. Operations Alerts"
        
        $issueVMs = $allVMs | Where-Object { $_.Runtime.ConnectionState -match "orphaned|inaccessible|invalid" -or $_.Runtime.ConsolidationNeeded -eq $true }
        $ovmRows = @()
        foreach ($ivm in $issueVMs) { 
            if ($ivm.Runtime.ConnectionState -match "orphaned|inaccessible|invalid") { $GlobalAlerts += "🚨 <b>[$vc]</b> $($ivm.Name) 고아 상태 ($($ivm.Runtime.ConnectionState))" }
            if ($ivm.Runtime.ConsolidationNeeded -eq $true) { $GlobalAlerts += "⚠️ <b>[$vc]</b> $($ivm.Name) 디스크 Consolidation 필요" }
            
            $stateBadge = if ($ivm.Runtime.ConnectionState -ne "connected") { (Get-BadgeHtml $ivm.Runtime.ConnectionState $EC.red) } else { "" }
            $consBadge  = if ($ivm.Runtime.ConsolidationNeeded -eq $true) { (Get-BadgeHtml "Needs Consolidation" $EC.yellow) } else { "" }
            $ovmRows += ,@("<b style='color:$($EC.red)'>$($ivm.Name)</b>", "$stateBadge $consBadge")
        }
        if ($ovmRows) { 
            $HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>📌 Orphaned, Inaccessible or Needs Consolidation VMs</div>"
            $HtmlDetails += Build-EmailTable -Headers @("VM Name", "Issue Detected") -Rows $ovmRows 
        }

        # Recent Events (Get-VIEvent 장애/경고 통합)
        $alertEvents = Get-VIEvent -Server $conn -Start $StartTime -MaxSamples 2000 | Where-Object { 
            $_.FullFormattedMessage -match "Error|Fail|HA|Warning|Alarm|down|die" -or 
            $_ -is [VMware.Vim.VmFailedToPowerOnEvent] -or
            $_ -is [VMware.Vim.EventEx]
        } | Sort-Object CreatedTime -Descending | Select-Object -First 30

        $evtRows = @()
        foreach ($e in $alertEvents) {
            $timeStr = $e.CreatedTime.ToUniversalTime().ToString('MM-dd HH:mm') + " UTC"
            $targetStr = if ($e.Vm) { $e.Vm.Name } elseif ($e.Host) { $e.Host.Name } elseif ($e.ObjectName) { $e.ObjectName } else { "System" }
            
            $typeStr = $e.GetType().Name.Replace("Event","")
            $badge = if ($typeStr -match "Error|Fail") { (Get-BadgeHtml "Error" $EC.red) } 
                     elseif ($typeStr -match "Warning") { (Get-BadgeHtml "Warning" $EC.yellow) } 
                     elseif ($e.FullFormattedMessage -match "vSphere HA|High Availability") { (Get-BadgeHtml "HA" $EC.pink) } 
                     else { (Get-BadgeHtml "Alert" $EC.muted) }
            
            $evtRows += ,@($timeStr, $badge, "<b>$targetStr</b>", $e.FullFormattedMessage)
        }
        $HtmlDetails += "<div style='margin-bottom:6px;font-weight:bold;'>📌 Recent Critical Events & Alerts (Top 30)</div>"
        $HtmlDetails += Build-EmailTable -Headers @("Time (UTC)", "Type", "Target", "Message / Details") -Rows $evtRows -EmptyMsg "최근 24시간 내 감지된 경고 및 장애 이벤트가 없습니다. ✔️"

    } catch {
        Write-DetailedLog "$vc 처리 중 오류 발생: $($_.Exception.Message)" "ERROR"
        $HtmlDetails += "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='margin-bottom:20px;'><tr><td style='background-color:#fee2e2;color:#991b1b;padding:15px;font-weight:bold;border-radius:5px;'>❌ $vc 점검 실패: $($_.Exception.Message)</td></tr></table>"
        $GlobalAlerts += "🚨 <b>[$vc]</b> 연결 두절 또는 데이터 수집 실패"
    } finally {
        if ($conn) { Disconnect-VIServer -Server $conn -Confirm:$false -WarningAction SilentlyContinue | Out-Null }
    }
}

try { $CurrentVMHistory | ConvertTo-Json -Depth 3 | Out-File $HistoryFilePath -Encoding UTF8 } catch {}

# ------------------------------------------------------------
# 5. HTML 최종 조합 및 타임스탬프 파일 생성
# ------------------------------------------------------------
$UtcHeaderTime = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

$HtmlHeader = @"
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:$($EC.bg);padding:20px;font-family:Arial,sans-serif;">
<tr><td align="center">
<table width="1000" cellpadding="0" cellspacing="0" border="0" style="background-color:$($EC.surface);border:1px solid $($EC.border);border-radius:8px;width:1000px;">
<tr><td style="background-color:$($EC.th_bg);padding:25px;border-radius:8px 8px 0 0;">
    <h1 style="color:#ffffff;margin:0;font-size:22px;">🛡️ VMware Infrastructure Daily Health Check</h1>
    <p style="color:#94a3b8;margin:5px 0 0 0;font-size:12px;">Standard Generation Time: $UtcHeaderTime UTC | Target Period: Last 24 Hours</p>
</td></tr>
<tr><td style="padding:25px;">
"@

$HtmlAttention = "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='margin-bottom:10px;'><tr><td style='font-size:16px;font-weight:bold;color:$($EC.red);font-family:Arial,sans-serif;border-bottom:2px solid #fca5a5;padding-bottom:5px;'>📢 전체 인프라 주요 주의사항 (Attention)</td></tr></table>"

if ($GlobalAlerts.Count -gt 0) {
    $HtmlAttention += "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='background-color:#fef2f2;border:1px solid #fecaca;border-radius:8px;margin-bottom:30px;'><tr><td style='padding:15px;'>"
    $HtmlAttention += "<ul style='margin:0;padding-left:20px;color:#991b1b;font-size:13px;line-height:1.6;font-family:Arial,sans-serif;'>"
    foreach ($alert in $GlobalAlerts) { $HtmlAttention += "<li>$alert</li>" }
    $HtmlAttention += "</ul></td></tr></table>"
} else {
    $HtmlAttention += "<table width='100%' cellpadding='0' cellspacing='0' border='0' style='background-color:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;margin-bottom:30px;'><tr><td style='padding:15px;color:#166534;font-size:13px;font-weight:bold;font-family:Arial,sans-serif;'>"
    $HtmlAttention += "✅ 전체 데이터센터 인프라에 장애 징후 및 집적도 과할당이 없습니다."
    $HtmlAttention += "</td></tr></table>"
}

$HtmlTrend = Build-SectionHeader "📊" "1. 전일 대비 VM 증감 추적 (VM Inventory Trends)"
$HtmlTrend += Build-EmailTable -Headers @("vCenter FQDN", "전일 VM 수", "금일 VM 수", "증감 (Diff)", "신규 생성 VM (+)", "삭제/미검출 VM (-)") -Rows $VmTrendRows

$HtmlFooter = "</td></tr></table></td></tr></table>" 
$FinalHtml = "<!DOCTYPE html><html lang='ko'><head><meta charset='UTF-8'></head><body>" + $HtmlHeader + $HtmlAttention + $HtmlTrend + $HtmlDetails + $HtmlFooter + "</body></html>"

# 💡 동적 파일명 생성 (DailyHealthCheck_YYYYMMDD_HHMM.html)
$TimeStamp = Get-Date -Format 'yyyyMMdd_HHmm'
$TargetFileName = "DailyHealthCheck_$TimeStamp.html"
$LocalHtmlPath = "$PSScriptRoot\$TargetFileName"

# 1. 로컬에 오늘 날짜 타임스탬프 파일 저장
$FinalHtml | Out-File -FilePath $LocalHtmlPath -Encoding UTF8 -Force

# 2. GitHub Pages용 호스팅 index.html 동시 갱신
$IndexHtmlPath = "$PSScriptRoot\index.html"
$FinalHtml | Out-File -FilePath $IndexHtmlPath -Encoding UTF8 -Force

Write-DetailedLog "로컬 HTML 리포트 생성 완료: $TargetFileName" "SUCCESS"

# ------------------------------------------------------------
# 6. 디테일 진단 로깅 기능이 탑재된 Pure PowerShell SMTP 이메일 엔진
# ------------------------------------------------------------
Write-DetailedLog "Gmail SMTP 진단 및 이메일 발송 세션을 시작합니다..." "INFO"

$cleanPass = if ($SmtpPass) { $SmtpPass.Replace(" ", "") } else { "" }

try {
    # 1. SMTP 서버 포트 도달 가능성 사전 테스트 (PowerShell 변수 파싱 에러 방지 구문 적용)
    $socket = New-Object System.Net.Sockets.TcpClient
    $connectTask = $socket.ConnectAsync($SmtpServer, $SmtpPort)
    if (-not $connectTask.Wait(3000)) {
        throw "SMTP 서버 통신 타임아웃: ${SmtpServer}:${SmtpPort} 포트에 접근할 수 없습니다. (방화벽/백신 블로킹 확인 필요)"
    }
    $socket.Close()
    Write-DetailedLog "SMTP 포트 연결 확인 완료 (${SmtpServer}:${SmtpPort})" "INFO"

    # 2. MailMessage 및 Client 객체 생성
    $mailMsg = New-Object System.Net.Mail.MailMessage
    $mailMsg.From = New-Object System.Net.Mail.MailAddress($From)
    $mailMsg.To.Add($From) # 발신자 본인을 To로 설정
    
    # BCC 숨은 참조 다중 수신자 등록
    foreach ($bcc in $BccRecipients) {
        if ($bcc) { $mailMsg.Bcc.Add($bcc) }
    }
    
    $mailMsg.Subject = $Subject
    $mailMsg.Body = $FinalHtml
    $mailMsg.IsBodyHtml = $true
    $mailMsg.BodyEncoding = [System.Text.Encoding]::UTF8
    $mailMsg.SubjectEncoding = [System.Text.Encoding]::UTF8

    $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $smtpClient.EnableSsl = $true
    $smtpClient.Timeout = 15000 # 15초 세션 타임아웃
    $smtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $cleanPass)

    # 3. 메일 발송 실행
    $smtpClient.Send($mailMsg)
    $mailMsg.Dispose()

    Write-DetailedLog "이메일 발송 성공! ($($BccRecipients.Count)명 수신자에게 BCC 발송됨)" "SUCCESS"

} catch {
    $errObj = $_.Exception
    Write-DetailedLog "이메일 발송 중 장애가 발생했습니다." "ERROR"
    Write-DetailedLog "  - 에러 원인: $($errObj.Message)" "ERROR"
    
    if ($errObj.InnerException) {
        Write-DetailedLog "  - 상세 예외: $($errObj.InnerException.Message)" "ERROR"
    }

    if ($errObj.Message -match "5.7.0|5.7.8|Authentication Required|BadCredentials") {
        Write-DetailedLog "  👉 [진단 힌트] Gmail 계정 인증 실패입니다. Google 계정 설정에서 '2단계 인증'이 켜져 있는지, 발급받은 16자리 '앱 비밀번호'가 올바른지 확인하세요." "WARN"
    } elseif ($errObj.Message -match "time out|타임아웃|Socket") {
        Write-DetailedLog "  👉 [진단 힌트] 네트워크 타임아웃입니다. 서버 또는 PC의 방화벽/백신에서 587 포트 아웃바운드 통신이 허용되어 있는지 확인하세요." "WARN"
    }
}

# ------------------------------------------------------------
# 7. GitHub REST API Dynamic 업로드 (오늘 날짜 타임스탬프 파일명 적용)
# ------------------------------------------------------------
Write-DetailedLog "GitHub REST API를 통한 레포지토리 업로드를 시작합니다..." "INFO"

if ($GitHubToken -and $GitHubRepo -and $GitHubToken -notmatch "ghp_xxxx") {
    try {
        # 💡 요청사항 반영: 오늘 날짜 타임스탬프 파일명으로 GitHub에 직접 등록
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

        # 1. 동일한 파일명의 SHA 커밋 해시 수집 (덮어쓰기 대비)
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
            $reader = New-Object System.IO.StreamReader($ghErr.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            Write-DetailedLog "  👉 GitHub 상세 응답: $responseBody" "WARN"
        }
    }
} else {
    Write-DetailedLog "GitHubToken 또는 GitHubRepo 설정이 유효하지 않아 GitHub 업로드를 건너뜁니다." "WARN"
}

Write-DetailedLog "모든 작업이 완료되었습니다." "INFO"
