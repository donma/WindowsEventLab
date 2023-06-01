
Param (
    # 查詢區間之起始時間(預設最近12小時內)
    [DateTime]$start = (Get-Date).AddHours(-12)
)
$wp = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-Not $wp.IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
	Write-Host "*** 請使用系統管理員權限執行 ***" -ForegroundColor Red
	return
}
# 登入類別對照表
$LogonTypeTexts = @(
    'NA','NA',
    'Interactive', #2
    'Network','Batch','Service','NA','Unlock','NetworkClearText',
    'NewCredentials','RemoteInteractive','CachedInteractive'
)
# 計算起始時間距今的毫秒數
$timeDiff = (New-TimeSpan -Start $start -End (Get-Date)).TotalMilliseconds
# 限定 4624(登入成功)、4625(登入失敗)
$xpath = @"
*[  
    System[
        (EventID=4624 or EventID=4625) and
        TimeCreated[timediff(`@SystemTime) <= $timeDiff]
    ] and 
    EventData[
        Data[@Name='IpAddress'] != '-' and
        Data[@Name='IpAddress'] != '::1' and
        Data[@Name='IpAddress'] != '127.0.0.1'
    ]
]
"@
# 加上 SilentlyContinue 防止查無資料時噴錯 No events were found that match the specified selection criteria.
Get-WinEvent -LogName 'Security' -FilterXPath $xpath -ErrorAction SilentlyContinue | ForEach-Object {
    $xml = [xml]$_.ToXml() # 將事件記錄轉成 XML
    $d = @{} # 建立 Hashtable 放 EventData.Data 中的客製屬性
    @($xml.Event.EventData.Data) | ForEach-Object {
        $d[$_.name] = $_.'#text'
    }
    if ($_.ID -eq 4624) { $action = '登入成功'  }
    elseif ($_.ID -eq 4625) { $action = '登入失敗' }
    $logonType = ''
    if ($d.LogonType -gt 1) {
        $logonType = $LogonTypeTexts[$d.LogonType]
    }
    [PSCustomObject]@{
        Action = $action;
        Time = $_.TimeCreated.ToString("yyyy/MM/dd HH:mm:ss");
        Id = $_.ID;
        TargetAccount = "$($d.TargetDomainName)\$($d.TargetUserName)"; # 登入帳號
        Socket = "$($d['IpAddress']):$($d['IpPort'])"; # IP 來源
        LogonType = $logonType +':'+ $d.LogonType ; # 登入類別
        LogonProcess = $d.LogonProcessName; # 程序名稱
        AuthPkgName = $d.AuthenticationPackageName; # 驗證模組
        SubjectLogonId=$d.SubjectLogonId
    }
} | ConvertTo-Json

#Soruce From 黑暗執行緒: https://blog.darkthread.net/blog/ps-list-logon-events/
#我只是做了部分修改

