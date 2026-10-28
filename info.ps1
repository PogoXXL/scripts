#requires -Version 5.1
<#
.SYNOPSIS
    Собирает отчёт о компьютере для диагностики.

.DESCRIPTION
    Показывает модель, комплектующие, версию Windows, диски, сеть, батарею,
    Secure Boot/TPM и дополнительные сведения. Работает без администратора,
    но часть данных (TPM, Defender, BitLocker) может быть недоступна.

.PARAMETER Json
    Вывести структурированный JSON вместо текста.

.PARAMETER OutFile
    Сохранить отчёт в файл (UTF-8). Формат определяется ключом -Json.

.PARAMETER NoPersonal
    Скрыть серийные номера, имя пользователя и MAC-адреса.

.PARAMETER Events
    Добавить сводку ошибок System/Application за последние 7 дней.

.EXAMPLE
    .\Get-PcReport.ps1

.EXAMPLE
    .\Get-PcReport.ps1 -OutFile C:\Temp\pc-report.txt

.EXAMPLE
    .\Get-PcReport.ps1 -Json -OutFile C:\Temp\pc-report.json
#>
[CmdletBinding()]
param(
    [switch]$Json,
    [string]$OutFile,
    [switch]$NoPersonal,
    [switch]$Events
)

$ErrorActionPreference = 'Continue'
$script:Lines = New-Object 'System.Collections.Generic.List[string]'

# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-RegValue {
    param([string]$Path, [string]$Name, $Default = $null)
    try {
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
        $value = $item.$Name
        if ($null -eq $value) { return $Default }
        return $value
    } catch {
        return $Default
    }
}

function Get-CimData {
    param(
        [string]$Class,
        [string]$Namespace = 'root/cimv2',
        [string]$Filter
    )
    try {
        if ($Filter) {
            return @(Get-CimInstance -ClassName $Class -Namespace $Namespace -Filter $Filter -ErrorAction Stop)
        }
        return @(Get-CimInstance -ClassName $Class -Namespace $Namespace -ErrorAction Stop)
    } catch {
        return @()
    }
}

function Format-DateValue {
    param($Value)
    if ($null -eq $Value) { return $null }
    try { return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm') } catch { return [string]$Value }
}

function Format-Bytes {
    param($Bytes)
    if ($null -eq $Bytes -or "$Bytes" -eq '') { return $null }
    $value = [double]$Bytes
    if ($value -ge 1TB) { return ('{0:N2} TB' -f ($value / 1TB)) }
    if ($value -ge 1GB) { return ('{0:N1} GB' -f ($value / 1GB)) }
    if ($value -ge 1MB) { return ('{0:N1} MB' -f ($value / 1MB)) }
    return ('{0:N0} KB' -f ($value / 1KB))
}

function Format-LinkSpeed {
    param($BitsPerSecond)
    if ($null -eq $BitsPerSecond -or "$BitsPerSecond".Trim() -eq '') { return $null }

    # Некоторые CIM-провайдеры отдают уже готовую строку ("1 Gbps" /
    # "100 Мбит/с"). Не пытаемся делить её на число.
    if ($BitsPerSecond -is [string]) {
        $text = $BitsPerSecond.Trim()
        if ($text -match '(?i)(Гбит|Мбит|Кбит|Gbps|Mbps|Kbps)') { return $text }
        $parsed = 0.0
        if (-not [double]::TryParse($text, [ref]$parsed)) { return $text }
        $BitsPerSecond = $parsed
    }

    $value = [double]$BitsPerSecond
    if ($value -le 0) { return $null }
    if ($value -ge 1e9) { return ('{0:N1} Гбит/с' -f ($value / 1e9)) }
    return ('{0:N0} Мбит/с' -f ($value / 1e6))
}

function Format-Uptime {
    param($Span)
    if ($null -eq $Span) { return $null }
    if ($Span.TotalSeconds -lt 0) { return $null }
    return ('{0} д {1} ч {2} мин' -f [int]$Span.TotalDays, $Span.Hours, $Span.Minutes)
}

function Get-MemoryTypeName {
    param($Code)
    $map = @{
        20 = 'DDR'; 21 = 'DDR2'; 24 = 'DDR3'; 26 = 'DDR4'; 34 = 'DDR5';
        18 = 'DDR'; 19 = 'DDR2'; 22 = 'DDR2'; 25 = 'DDR3'; 27 = 'DDR3'
    }
    if ($map.ContainsKey([int]$Code)) { return $map[[int]$Code] }
    return $null
}

function Get-FormFactorName {
    param($Code)
    $map = @{ 8 = 'DIMM'; 12 = 'SODIMM'; 13 = 'SODIMM'; 11 = 'DIMM' }
    if ($map.ContainsKey([int]$Code)) { return $map[[int]$Code] }
    return $null
}

function Get-SystemTypeName {
    param($PcType, $Chassis)
    if ($Chassis) {
        foreach ($item in @($Chassis)) {
            $value = [int]$item
            if ($value -in 8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32) { return 'Ноутбук / планшет' }
            if ($value -in 3, 4, 5, 6, 7, 13, 15, 16, 35, 36) { return 'Настольный ПК' }
            if ($value -in 23, 28) { return 'Сервер' }
        }
    }
    switch ([int]$PcType) {
        1 { return 'Настольный ПК' }
        2 { return 'Ноутбук' }
        3 { return 'Рабочая станция' }
        4 { return 'Сервер' }
        5 { return 'Сервер' }
        6 { return 'Прибор' }
        7 { return 'Сервер' }
        default { return $null }
    }
}

function Add-Line {
    param([string]$Text = '')
    $script:Lines.Add($Text)
}

function Add-Section {
    param([string]$Title)
    Add-Line ''
    Add-Line ('=' * 78)
    Add-Line (' ' + $Title.ToUpperInvariant())
    Add-Line ('=' * 78)
}

function Add-Value {
    param([string]$Label, $Value, [int]$Width = 24)
    if ($null -eq $Value -or "$Value".Trim() -eq '') { $Value = 'n/a' }
    Add-Line ('  ' + $Label.PadRight($Width) + ': ' + $Value)
}

function Add-Table {
    param(
        [object[]]$Rows,
        [string[]]$Columns,
        [string[]]$Headers
    )
    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        Add-Line '  (нет данных)'
        return
    }
    if (-not $Headers) { $Headers = $Columns }

    $matrix = @()
    foreach ($row in $Rows) {
        $cells = @()
        foreach ($column in $Columns) {
            $value = Get-Prop $row $column
            if ($null -eq $value -or "$value".Trim() -eq '') {
                $cells += '-'
            } else {
                $cells += ([string]$value).Trim()
            }
        }
        $matrix += ,$cells
    }

    $widths = @()
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        $max = $Headers[$i].Length
        foreach ($cells in $matrix) {
            if ($cells[$i].Length -gt $max) { $max = $cells[$i].Length }
        }
        $widths += [Math]::Min($max, 44)
    }

    $render = {
        param($Cells)
        $parts = @()
        for ($i = 0; $i -lt $Cells.Count; $i++) {
            $text = [string]$Cells[$i]
            if ($text.Length -gt $widths[$i]) {
                $text = $text.Substring(0, [Math]::Max(1, $widths[$i] - 3)) + '...'
            }
            $parts += $text.PadRight($widths[$i])
        }
        return '  ' + ($parts -join '  ')
    }

    Add-Line (& $render -Cells $Headers)
    Add-Line ('  ' + (($widths | ForEach-Object { '-' * $_ }) -join '  '))
    foreach ($cells in $matrix) {
        Add-Line (& $render -Cells $cells)
    }
}

# ---------------------------------------------------------------------------
# Сбор данных
# ---------------------------------------------------------------------------

$now = Get-Date
$biosRegPath = 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS'
$osRegPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$cpuRegPath = 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0'

$computer = Get-CimData -Class 'Win32_ComputerSystem' | Select-Object -First 1
$bios = Get-CimData -Class 'Win32_BIOS' | Select-Object -First 1
$board = Get-CimData -Class 'Win32_BaseBoard' | Select-Object -First 1
$enclosure = Get-CimData -Class 'Win32_SystemEnclosure' | Select-Object -First 1
$os = Get-CimData -Class 'Win32_OperatingSystem' | Select-Object -First 1
$cpus = @(Get-CimData -Class 'Win32_Processor')
$modules = @(Get-CimData -Class 'Win32_PhysicalMemory')
$memoryArray = Get-CimData -Class 'Win32_PhysicalMemoryArray' | Select-Object -First 1

# --- Система ---------------------------------------------------------------
$manufacturer = Get-Prop $computer 'Manufacturer' (Get-RegValue $biosRegPath 'SystemManufacturer')
$model = Get-Prop $computer 'Model' (Get-RegValue $biosRegPath 'SystemProductName')
$serial = Get-Prop $bios 'SerialNumber'
if ($NoPersonal -and $serial) { $serial = 'скрыт' }
$biosVersion = Get-Prop $bios 'SMBIOSBIOSVersion' (Get-RegValue $biosRegPath 'BIOSVersion')
$biosDate = Format-DateValue (Get-Prop $bios 'ReleaseDate' (Get-RegValue $biosRegPath 'BIOSReleaseDate'))
$boardManufacturer = Get-Prop $board 'Manufacturer' (Get-RegValue $biosRegPath 'BaseBoardManufacturer')
$boardProduct = Get-Prop $board 'Product' (Get-RegValue $biosRegPath 'BaseBoardProduct')
$pcType = Get-SystemTypeName (Get-Prop $computer 'PCSystemType') (Get-Prop $enclosure 'ChassisTypes')
$domain = Get-Prop $computer 'Domain'
$partOfDomain = Get-Prop $computer 'PartOfDomain' $false
$userName = Get-Prop $computer 'UserName'
if (-not $userName) { $userName = "$env:USERDOMAIN\$env:USERNAME" }
if ($NoPersonal -and $userName) { $userName = 'скрыт' }

$biosText = $null
if ($biosVersion) {
    $biosText = $biosVersion.Trim()
    if ($biosDate) { $biosText += " ($biosDate)" }
}
$boardText = $null
if ($boardProduct) {
    $boardText = ("{0} {1}" -f $boardManufacturer, $boardProduct).Trim()
}

$computerInfo = [ordered]@{
    name        = $env:COMPUTERNAME
    manufacturer = if ($manufacturer) { $manufacturer.Trim() } else { $null }
    model       = if ($model) { $model.Trim() } else { $null }
    type        = $pcType
    serial      = if ($serial) { $serial.Trim() } else { $null }
    domain      = $domain
    part_of_domain = [bool]$partOfDomain
    user        = $userName
    bios        = $biosText
    board       = $boardText
}

# --- ОС --------------------------------------------------------------------
$productName = Get-Prop $os 'Caption' (Get-RegValue $osRegPath 'ProductName' 'Windows')
$displayVersion = Get-RegValue $osRegPath 'DisplayVersion'
$build = Get-Prop $os 'BuildNumber' (Get-RegValue $osRegPath 'CurrentBuild')
$ubr = Get-RegValue $osRegPath 'UBR'
$architecture = Get-Prop $os 'OSArchitecture' $env:PROCESSOR_ARCHITECTURE
$installDate = Get-Prop $os 'InstallDate'
if (-not $installDate) {
    $installUnix = Get-RegValue $osRegPath 'InstallDate'
    if ($installUnix -is [int] -and $installUnix -gt 0) {
        try { $installDate = [DateTimeOffset]::FromUnixTimeSeconds($installUnix).LocalDateTime } catch { $installDate = $null }
    }
}
$lastBoot = Get-Prop $os 'LastBootUpTime'
$uptime = $null
if ($lastBoot) {
    try { $uptime = (Get-Date) - ([datetime]$lastBoot) } catch { $uptime = $null }
}
$timeZone = try { (Get-TimeZone).DisplayName } catch { $null }

$activation = $null
$licenses = @(Get-CimData -Class 'SoftwareLicensingProduct' -Filter "PartialProductKey is not null AND Name like 'Windows%'")
foreach ($license in $licenses) {
    $status = Get-Prop $license 'LicenseStatus'
    if ($status -eq 1) { $activation = 'активирована'; break }
}
if (-not $activation -and $licenses.Count -gt 0) {
    $activation = 'не активирована (код ' + (Get-Prop $licenses[0] 'LicenseStatus' '?') + ')'
}

$osInfo = [ordered]@{
    name        = if ($productName) { $productName.Trim() } else { $null }
    display_version = $displayVersion
    build       = if ($build) { if ($ubr) { "$build.$ubr" } else { "$build" } } else { $null }
    architecture = $architecture
    installed   = Format-DateValue $installDate
    last_boot   = Format-DateValue $lastBoot
    uptime      = Format-Uptime $uptime
    timezone    = $timeZone
    activation  = $activation
}

# --- Процессор -------------------------------------------------------------
$cpu = $cpus | Select-Object -First 1
$cpuName = Get-Prop $cpu 'Name' (Get-RegValue $cpuRegPath 'ProcessorNameString')
$cores = $null
$threads = $null
if ($cpus.Count -gt 0) {
    $cores = ($cpus | Measure-Object -Property NumberOfCores -Sum).Sum
    $threads = ($cpus | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
}
if (-not $threads -and $env:NUMBER_OF_PROCESSORS) { $threads = [int]$env:NUMBER_OF_PROCESSORS }
$maxClock = Get-Prop $cpu 'MaxClockSpeed' (Get-RegValue $cpuRegPath '~MHz')
$socket = Get-Prop $cpu 'SocketDesignation'

$cpuInfo = [ordered]@{
    name        = if ($cpuName) { $cpuName.Trim() } else { $null }
    cores       = $cores
    threads     = $threads
    max_clock_mhz = $maxClock
    socket      = $socket
    virtualization = Get-Prop $cpu 'VirtualizationFirmwareEnabled'
}

# --- Память ----------------------------------------------------------------
$totalRam = ($modules | Measure-Object -Property Capacity -Sum).Sum
if (-not $totalRam) {
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
        $totalRam = (New-Object Microsoft.VisualBasic.Devices.ComputerInfo).TotalPhysicalMemory
    } catch {
        $totalRam = $null
    }
}
$slotCount = Get-Prop $memoryArray 'MemoryDevices'
if (-not $slotCount -and $modules.Count -gt 0) { $slotCount = $modules.Count }
$slotsFree = $null
if ($slotCount) { $slotsFree = [Math]::Max(0, [int]$slotCount - $modules.Count) }
$moduleRows = @()
foreach ($module in $modules) {
    $sizeBytes = Get-Prop $module 'Capacity'
    $speed = Get-Prop $module 'ConfiguredClockSpeed' (Get-Prop $module 'Speed')
    $typeName = Get-MemoryTypeName (Get-Prop $module 'SMBIOSMemoryType')
    $formFactor = Get-FormFactorName (Get-Prop $module 'FormFactor')
    $typeText = (@($typeName, $formFactor) | Where-Object { $_ }) -join ' '
    $moduleRows += [pscustomobject]@{
        slot         = Get-Prop $module 'DeviceLocator' (Get-Prop $module 'BankLabel')
        size_bytes   = $sizeBytes
        size         = Format-Bytes $sizeBytes
        speed_mhz    = $speed
        type         = $typeText
        manufacturer = (Get-Prop $module 'Manufacturer' '').Trim()
        part_number  = (Get-Prop $module 'PartNumber' '').Trim()
    }
}
$memoryInfo = [ordered]@{
    total_bytes = $totalRam
    total       = Format-Bytes $totalRam
    slots_total = $slotCount
    slots_used  = $modules.Count
    slots_free  = $slotsFree
    modules     = $moduleRows
}

# --- Видеокарты ------------------------------------------------------------
$gpuRows = @()
$gpus = @(Get-CimData -Class 'Win32_VideoController')
if ($gpus.Count -gt 0) {
    foreach ($gpu in $gpus) {
        $vram = Get-Prop $gpu 'AdapterRAM'
        if ($vram -and [int64]$vram -lt 0) { $vram = $null }
        $resolution = $null
        $width = Get-Prop $gpu 'CurrentHorizontalResolution'
        $height = Get-Prop $gpu 'CurrentVerticalResolution'
        $refresh = Get-Prop $gpu 'CurrentRefreshRate'
        if ($width -and $height) {
            $resolution = "$width x $height"
            if ($refresh) { $resolution += " @ $refresh Hz" }
        }
        $gpuRows += [pscustomobject]@{
            name        = Get-Prop $gpu 'Name'
            vram_bytes  = $vram
            vram        = Format-Bytes $vram
            driver      = Get-Prop $gpu 'DriverVersion'
            driver_date = Format-DateValue (Get-Prop $gpu 'DriverDate')
            resolution  = $resolution
        }
    }
} else {
    # Запасной путь: реестр драйверов дисплея.
    $displayClass = 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}'
    foreach ($key in @(Get-ChildItem -LiteralPath $displayClass -ErrorAction SilentlyContinue)) {
        if ($key.PSChildName -notmatch '^\d{4}$') { continue }
        $values = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
        if (-not $values.DriverDesc) { continue }
        $gpuRows += [pscustomobject]@{
            name        = $values.DriverDesc
            vram_bytes  = $null
            vram        = $null
            driver      = $values.DriverVersion
            driver_date = $values.DriverDate
            resolution  = $null
        }
    }
}

# --- Диски -----------------------------------------------------------------
$diskRows = @()
try {
    $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop)
} catch {
    $physicalDisks = @()
}
if ($physicalDisks.Count -gt 0) {
    foreach ($disk in $physicalDisks) {
        $diskRows += [pscustomobject]@{
            model      = Get-Prop $disk 'FriendlyName'
            size_bytes = Get-Prop $disk 'Size'
            size       = Format-Bytes (Get-Prop $disk 'Size')
            media      = Get-Prop $disk 'MediaType'
            bus        = Get-Prop $disk 'BusType'
            health     = Get-Prop $disk 'HealthStatus'
            serial     = if ($NoPersonal) { 'скрыт' } else { Get-Prop $disk 'SerialNumber' }
        }
    }
} else {
    foreach ($disk in @(Get-CimData -Class 'Win32_DiskDrive')) {
        $diskRows += [pscustomobject]@{
            model      = Get-Prop $disk 'Model'
            size_bytes = Get-Prop $disk 'Size'
            size       = Format-Bytes (Get-Prop $disk 'Size')
            media      = Get-Prop $disk 'MediaType'
            bus        = Get-Prop $disk 'InterfaceType'
            health     = Get-Prop $disk 'Status'
            serial     = if ($NoPersonal) { 'скрыт' } else { (Get-Prop $disk 'SerialNumber' '').Trim() }
        }
    }
}

$volumeRows = @()
try {
    $volumes = @(Get-Volume -ErrorAction Stop)
} catch {
    $volumes = @()
}
if ($volumes.Count -gt 0) {
    foreach ($volume in $volumes) {
        if (-not $volume.DriveLetter -and -not $volume.FileSystemLabel) { continue }
        if ($volume.DriveType -eq 'Unknown' -and -not $volume.Size) { continue }
        $sizeBytes = Get-Prop $volume 'Size'
        $freeBytes = Get-Prop $volume 'SizeRemaining'
        $freePct = $null
        if ($sizeBytes -and $freeBytes -ne $null) { $freePct = [Math]::Round(($freeBytes / $sizeBytes) * 100, 1) }
        $drive = if ($volume.DriveLetter) { "$($volume.DriveLetter):" } else { '—' }
        $volumeRows += [pscustomobject]@{
            drive      = $drive
            label      = Get-Prop $volume 'FileSystemLabel'
            fs         = Get-Prop $volume 'FileSystem'
            type       = Get-Prop $volume 'DriveType'
            size_bytes = $sizeBytes
            size       = Format-Bytes $sizeBytes
            free_bytes = $freeBytes
            free       = Format-Bytes $freeBytes
            free_pct   = if ($freePct -ne $null) { "$freePct%" } else { $null }
        }
    }
} else {
    # Запасной путь: логические диски через PSDrive.
    foreach ($drive in @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue)) {
        if ($drive.Name -notmatch '^[A-Za-z]$') { continue }
        if ($null -eq $drive.Free -or $null -eq $drive.Used) { continue }
        $sizeBytes = [double]$drive.Used + [double]$drive.Free
        $freePct = if ($sizeBytes -gt 0) { [Math]::Round(($drive.Free / $sizeBytes) * 100, 1) } else { $null }
        $volumeRows += [pscustomobject]@{
            drive      = "$($drive.Name):"
            label      = $null
            fs         = $null
            type       = 'Fixed'
            size_bytes = $sizeBytes
            size       = Format-Bytes $sizeBytes
            free_bytes = $drive.Free
            free       = Format-Bytes $drive.Free
            free_pct   = if ($freePct -ne $null) { "$freePct%" } else { $null }
        }
    }
}
$volumeRows = @($volumeRows | Sort-Object drive)

# --- Сеть ------------------------------------------------------------------
$networkRows = @()
try {
    $netConfigs = @(Get-NetIPConfiguration -ErrorAction Stop)
} catch {
    $netConfigs = @()
}
if ($netConfigs.Count -gt 0) {
    foreach ($config in $netConfigs) {
        $adapter = Get-Prop $config 'NetAdapter'
        $ipv4 = @(Get-Prop $config 'IPv4Address' @() | ForEach-Object { Get-Prop $_ 'IPAddress' })
        $ipv6 = @(Get-Prop $config 'IPv6Address' @() | ForEach-Object { Get-Prop $_ 'IPAddress' })
        $ips = @($ipv4 + $ipv6 | Where-Object { $_ })
        $ips = @($ips | Where-Object { $_ -notmatch '^fe80:' })
        if ($ips.Count -eq 0) { $ips = @($ipv4 + $ipv6 | Where-Object { $_ }) }
        $ips = @($ips | Sort-Object { if ($_ -match ':') { 1 } else { 0 } })
        $gateway = @(Get-Prop $config 'IPv4DefaultGateway' @() | ForEach-Object { Get-Prop $_ 'NextHop' })
        $dns = @(Get-Prop $config 'DNSServer' @() | ForEach-Object { Get-Prop $_ 'ServerAddresses' } | ForEach-Object { $_ })
        $dns = @($dns | Where-Object { $_ -and $_ -notmatch '^(fe80|fec0):' } | Select-Object -First 3)
        $usableIps = @($ips | Where-Object { $_ -notmatch '^169\.254\.' })
        if ($usableIps.Count -eq 0 -and $gateway.Count -eq 0) { continue }
        if ($ips.Count -eq 0 -and $gateway.Count -eq 0) { continue }
        $networkRows += [pscustomobject]@{
            name        = Get-Prop $config 'InterfaceAlias'
            description = Get-Prop $config 'InterfaceDescription'
            ip          = ($ips -join ', ')
            gateway     = ($gateway -join ', ')
            dns         = (($dns | Where-Object { $_ }) -join ', ')
            speed_bps   = Get-Prop $adapter 'LinkSpeed'
            speed       = Format-LinkSpeed (Get-Prop $adapter 'LinkSpeed')
            mac         = if ($NoPersonal) { 'скрыт' } else { Get-Prop $adapter 'MacAddress' }
            status      = Get-Prop $adapter 'Status'
            dhcp        = $null
        }
    }
} else {
    # Запасной путь: .NET, без WMI.
    try {
        foreach ($nic in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
            if ($nic.OperationalStatus -ne 'Up') { continue }
            if ($nic.NetworkInterfaceType -eq 'Loopback') { continue }
            $properties = $nic.GetIPProperties()
            $allIps = @($properties.UnicastAddresses | ForEach-Object { $_.Address.IPAddressToString })
            $ips = @($allIps | Where-Object { $_ -notmatch '^fe80:' })
            if ($ips.Count -eq 0) { $ips = $allIps }
            $ips = @($ips | Sort-Object { if ($_ -match ':') { 1 } else { 0 } })
            $gateway = @($properties.GatewayAddresses | ForEach-Object { $_.Address.IPAddressToString })
            $dns = @($properties.DnsAddresses | ForEach-Object { $_.IPAddressToString } | Where-Object { $_ -notmatch '^(fe80|fec0):' } | Select-Object -First 3)
            $usableIps = @($ips | Where-Object { $_ -notmatch '^169\.254\.' })
            if ($usableIps.Count -eq 0 -and $gateway.Count -eq 0) { continue }
            $networkRows += [pscustomobject]@{
                name        = $nic.Name
                description = $nic.Description
                ip          = ($ips -join ', ')
                gateway     = ($gateway -join ', ')
                dns         = ($dns -join ', ')
                speed_bps   = $nic.Speed
                speed       = Format-LinkSpeed $nic.Speed
                mac         = if ($NoPersonal) { 'скрыт' } else { $nic.GetPhysicalAddress().ToString() }
                status      = [string]$nic.OperationalStatus
                dhcp        = $null
            }
        }
    } catch {
        $networkRows = @()
    }
}
$networkRows = @($networkRows | Select-Object -First 8)

# --- Батарея ---------------------------------------------------------------
$batteryInfo = $null
$battery = Get-CimData -Class 'Win32_Battery' | Select-Object -First 1
if ($battery) {
    $static = Get-CimData -Class 'BatteryStaticData' -Namespace 'root/wmi' | Select-Object -First 1
    $full = Get-CimData -Class 'BatteryFullChargedCapacity' -Namespace 'root/wmi' | Select-Object -First 1
    $status = Get-CimData -Class 'BatteryStatus' -Namespace 'root/wmi' | Select-Object -First 1
    $cycles = Get-CimData -Class 'BatteryCycleCount' -Namespace 'root/wmi' | Select-Object -First 1
    $design = Get-Prop $static 'DesignedCapacity'
    $fullCharge = Get-Prop $full 'FullChargedCapacity'
    $health = $null
    if ($design -and $fullCharge) { $health = [Math]::Round(($fullCharge / $design) * 100, 1) }
    $batteryStatus = Get-Prop $status 'BatteryStatus' (Get-Prop $battery 'BatteryStatus')
    $statusText = switch ([int]$batteryStatus) {
        1 { 'разряжается' }
        2 { 'от сети' }
        3 { 'полностью заряжена' }
        4 { 'низкий заряд' }
        5 { 'критический заряд' }
        6 { 'заряжается' }
        7 { 'заряжается' }
        8 { 'заряжается' }
        9 { 'заряжается' }
        11 { 'частично заряжена' }
        default { $null }
    }
    $batteryInfo = [ordered]@{
        name        = Get-Prop $battery 'Name'
        status      = $statusText
        charge_pct  = Get-Prop $status 'EstimatedChargeRemaining' (Get-Prop $battery 'EstimatedChargeRemaining')
        design_mwh  = $design
        full_mwh    = $fullCharge
        health_pct  = $health
        cycle_count = Get-Prop $cycles 'CycleCount'
    }
}

if (-not $pcType -and $batteryInfo) {
    $pcType = 'Ноутбук (определено по батарее)'
    $computerInfo['type'] = $pcType
}

# --- Безопасность ----------------------------------------------------------
$secureBoot = $null
try {
    $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop
} catch {
    $secureBootReg = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State' 'UEFISecureBootEnabled'
    if ($null -ne $secureBootReg) { $secureBoot = [bool]$secureBootReg }
}
$tpmPresent = $null
$tpmVersion = $null
try {
    $tpm = Get-Tpm -ErrorAction Stop
    $tpmPresent = Get-Prop $tpm 'TpmPresent'
    $tpmVersion = Get-Prop $tpm 'ManufacturerVersion'
} catch {
    $tpmCim = Get-CimData -Class 'Win32_Tpm' -Namespace 'root/cimv2/security/microsofttpm' | Select-Object -First 1
    if ($tpmCim) {
        $tpmPresent = $true
        $tpmVersion = Get-Prop $tpmCim 'SpecVersion'
    }
}
$defenderEnabled = $null
$defenderRealtime = $null
$defenderSignature = $null
try {
    $defender = Get-MpComputerStatus -ErrorAction Stop
    $defenderEnabled = Get-Prop $defender 'AntivirusEnabled'
    $defenderRealtime = Get-Prop $defender 'RealTimeProtectionEnabled'
    $defenderSignature = Get-Prop $defender 'AntivirusSignatureVersion'
} catch {
    $defenderEnabled = $null
}
$bitlocker = @()
try {
    $bitlocker = @(Get-BitLockerVolume -ErrorAction Stop | Where-Object { $_.VolumeType -ne 'Data' } | ForEach-Object {
        [pscustomobject]@{
            mount = $_.MountPoint
            status = [string]$_.ProtectionStatus
            percentage = $_.EncryptionPercentage
        }
    })
} catch {
    $bitlocker = @()
}

$securityInfo = [ordered]@{
    secure_boot        = $secureBoot
    tpm_present        = $tpmPresent
    tpm_version        = $tpmVersion
    defender_enabled   = $defenderEnabled
    defender_realtime  = $defenderRealtime
    defender_signature = $defenderSignature
    bitlocker          = $bitlocker
}

# --- Ожидающие перезагрузки ------------------------------------------------
$rebootReasons = @()
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
    $rebootReasons += 'обновления компонентов Windows'
}
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
    $rebootReasons += 'Windows Update'
}
$pendingRenames = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations'
if ($pendingRenames) { $rebootReasons += 'отложенные переименования файлов' }
$pendingReboot = [ordered]@{
    required = ($rebootReasons.Count -gt 0)
    reasons  = $rebootReasons
}

# --- Ошибки в журналах (опционально) ---------------------------------------
$eventRows = @()
if ($Events) {
    try {
        $start = (Get-Date).AddDays(-7)
        $logs = @(Get-WinEvent -FilterHashtable @{ LogName = @('System', 'Application'); Level = @(1, 2); StartTime = $start } -MaxEvents 300 -ErrorAction Stop)
        $eventRows = @($logs | Group-Object ProviderName, Id | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
            $first = $_.Group[0]
            [pscustomobject]@{
                provider = $first.ProviderName
                id       = $first.Id
                count    = $_.Count
                level    = $first.LevelDisplayName
                last     = Format-DateValue $first.TimeCreated
            }
        })
    } catch {
        $eventRows = @()
    }
}

# ---------------------------------------------------------------------------
# JSON
# ---------------------------------------------------------------------------
$report = [ordered]@{
    generated      = $now.ToString('yyyy-MM-dd HH:mm:ss')
    computer       = $computerInfo
    os             = $osInfo
    cpu            = $cpuInfo
    memory         = $memoryInfo
    gpus           = @($gpuRows)
    disks          = @($diskRows)
    volumes        = @($volumeRows)
    network        = @($networkRows)
    battery        = $batteryInfo
    security       = $securityInfo
    pending_reboot = $pendingReboot
    events         = @($eventRows)
}

if ($Json) {
    $jsonText = $report | ConvertTo-Json -Depth 10
    if ($OutFile) {
        $fullPath = [IO.Path]::GetFullPath($OutFile)
        $parent = Split-Path -Parent $fullPath
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        [IO.File]::WriteAllText($fullPath, $jsonText, (New-Object System.Text.UTF8Encoding($false)))
        Write-Host "JSON сохранён: $fullPath"
    } else {
        Write-Output $jsonText
    }
    return
}

# ---------------------------------------------------------------------------
# Текст
# ---------------------------------------------------------------------------
Add-Line ('=' * 78)
Add-Line (' ОТЧЁТ О КОМПЬЮТЕРЕ')
Add-Line (' ' + $env:COMPUTERNAME + '  |  ' + $now.ToString('yyyy-MM-dd HH:mm:ss'))
Add-Line ('=' * 78)

Add-Section 'Система'
Add-Value 'Компьютер' $env:COMPUTERNAME
Add-Value 'Тип' $pcType
Add-Value 'Производитель' $computerInfo.manufacturer
Add-Value 'Модель' $computerInfo.model
Add-Value 'Серийный номер' $computerInfo.serial
Add-Value 'Домен / группа' $(if ($domain) { $domain } else { 'рабочая группа' })
Add-Value 'Пользователь' $userName
Add-Value 'BIOS' $computerInfo.bios
Add-Value 'Материнская плата' $computerInfo.board

Add-Section 'Операционная система'
Add-Value 'Версия' $osInfo.name
Add-Value 'Выпуск' $osInfo.display_version
Add-Value 'Сборка' $osInfo.build
Add-Value 'Архитектура' $osInfo.architecture
Add-Value 'Установлена' $osInfo.installed
Add-Value 'Последняя загрузка' $osInfo.last_boot
Add-Value 'Время работы' $osInfo.uptime
Add-Value 'Часовой пояс' $osInfo.timezone
Add-Value 'Активация' $osInfo.activation

Add-Section 'Процессор'
Add-Value 'Модель' $cpuInfo.name
Add-Value 'Ядра / потоки' $(if ($cores -and $threads) { "$cores / $threads" } elseif ($cores) { "$cores" } elseif ($threads) { "н/д / $threads" } else { $null })
Add-Value 'Макс. частота' $(if ($maxClock) { "$maxClock MHz" } else { $null })
Add-Value 'Сокет' $socket
Add-Value 'Виртуализация' $(if ($cpuInfo.virtualization -eq $true) { 'включена' } elseif ($cpuInfo.virtualization -eq $false) { 'выключена' } else { $null })

Add-Section 'Память'
Add-Value 'Всего' $memoryInfo.total
Add-Value 'Слоты' $(if ($memoryInfo.slots_total) { "{0} (занято {1}, свободно {2})" -f $memoryInfo.slots_total, $memoryInfo.slots_used, $memoryInfo.slots_free } else { 'нет данных' })
if ($moduleRows.Count -gt 0) {
    Add-Table -Rows $moduleRows -Columns @('slot', 'size', 'speed_mhz', 'type', 'manufacturer', 'part_number') -Headers @('Слот', 'Объём', 'МГц', 'Тип', 'Производитель', 'Модель')
} else {
    Add-Line '  Модули: нет данных (WMI недоступен или нет прав)'
}

Add-Section 'Видео'
if ($gpuRows.Count -gt 0) {
    Add-Table -Rows $gpuRows -Columns @('name', 'vram', 'driver', 'driver_date', 'resolution') -Headers @('Модель', 'Видеопамять', 'Драйвер', 'Дата', 'Режим')
} else {
    Add-Line '  (нет данных)'
}

Add-Section 'Диски'
if ($diskRows.Count -gt 0) {
    Add-Table -Rows $diskRows -Columns @('model', 'size', 'media', 'bus', 'health') -Headers @('Модель', 'Объём', 'Тип', 'Шина', 'Состояние')
} else {
    Add-Line '  Физические диски: нет данных (WMI недоступен или нет прав)'
}
Add-Line ''
Add-Line '  Тома:'
if ($volumeRows.Count -gt 0) {
    Add-Table -Rows $volumeRows -Columns @('drive', 'label', 'fs', 'size', 'free', 'free_pct') -Headers @('Диск', 'Метка', 'ФС', 'Всего', 'Свободно', '%')
} else {
    Add-Line '  (нет данных)'
}

Add-Section 'Сеть'
if ($networkRows.Count -gt 0) {
    Add-Table -Rows $networkRows -Columns @('name', 'ip', 'gateway', 'dns', 'speed') -Headers @('Интерфейс', 'IP', 'Шлюз', 'DNS', 'Скорость')
} else {
    Add-Line '  (нет активных интерфейсов или нет данных)'
}

if ($batteryInfo) {
    Add-Section 'Питание'
    Add-Value 'Батарея' $batteryInfo.name
    Add-Value 'Состояние' $batteryInfo.status
    Add-Value 'Заряд' $(if ($batteryInfo.charge_pct -ne $null) { "$($batteryInfo.charge_pct)%" } else { $null })
    Add-Value 'Ёмкость (проект.)' $(if ($batteryInfo.design_mwh) { "$($batteryInfo.design_mwh) мВт·ч" } else { $null })
    Add-Value 'Ёмкость (факт.)' $(if ($batteryInfo.full_mwh) { "$($batteryInfo.full_mwh) мВт·ч" } else { $null })
    Add-Value 'Износ' $(if ($batteryInfo.health_pct -ne $null) { "$($batteryInfo.health_pct)% здоровья" } else { $null })
    Add-Value 'Циклы' $batteryInfo.cycle_count
}

Add-Section 'Безопасность'
Add-Value 'Secure Boot' $(if ($secureBoot -eq $true) { 'включена' } elseif ($secureBoot -eq $false) { 'выключена' } else { 'нет данных (нужны права/UEFI)' })
Add-Value 'TPM' $(if ($tpmPresent -eq $true) { "присутствует $(if ($tpmVersion) { $tpmVersion })" } elseif ($tpmPresent -eq $false) { 'отсутствует' } else { 'нет данных (нужны права)' })
Add-Value 'Defender' $(if ($defenderEnabled -eq $true) { "включён $(if ($defenderSignature) { ', базы ' + $defenderSignature })" } elseif ($defenderEnabled -eq $false) { 'выключен' } else { 'нет данных (нужны права)' })
Add-Value 'Защита (real-time)' $(if ($defenderRealtime -eq $true) { 'включена' } elseif ($defenderRealtime -eq $false) { 'выключена' } else { $null })
if ($bitlocker.Count -gt 0) {
    Add-Value 'BitLocker' (($bitlocker | ForEach-Object { "$($_.mount) $($_.status)" }) -join '; ')
} else {
    Add-Value 'BitLocker' 'нет данных (нужны права)'
}

Add-Section 'Перезагрузка'
Add-Value 'Требуется' $(if ($pendingReboot.required) { 'да' } else { 'нет' })
if ($pendingReboot.required) {
    Add-Value 'Причина' ($rebootReasons -join '; ')
}

if ($Events) {
    Add-Section 'Ошибки за 7 дней'
    if ($eventRows.Count -gt 0) {
        Add-Table -Rows $eventRows -Columns @('provider', 'id', 'count', 'level', 'last') -Headers @('Источник', 'ID', 'Кол-во', 'Уровень', 'Последняя')
    } else {
        Add-Line '  Ошибок не найдено или журнал недоступен.'
    }
}

Add-Line ''
Add-Line ('=' * 78)
Add-Line '  n/a — данные недоступны (нет прав, WMI или устройства).'
Add-Line '  Для TPM/Defender/BitLocker запустите PowerShell от администратора.'
Add-Line ('=' * 78)

$text = $script:Lines -join [Environment]::NewLine

if ($OutFile) {
    $fullPath = [IO.Path]::GetFullPath($OutFile)
    $parent = Split-Path -Parent $fullPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    [IO.File]::WriteAllText($fullPath, $text, (New-Object System.Text.UTF8Encoding($true)))
    Write-Host "Отчёт сохранён: $fullPath"
} else {
    Write-Output $text
}
