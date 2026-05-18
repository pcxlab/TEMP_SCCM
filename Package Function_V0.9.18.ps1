# ============================
# INIT
# ============================

if (-not (Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
}

$script:Config = @{
    LogPath    = "C:\Temp\PCX.log"
    RetryCount = 3
    RetryDelay = 5
}

# ============================
# LOGGING
# ============================

function Write-PCXLog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"

    try {
        Add-Content -Path $script:Config.LogPath -Value $entry
    }
    catch {}

    switch ($Level) {
        "ERROR" { Write-Host $entry -ForegroundColor Red }
        "WARN" { Write-Host $entry -ForegroundColor Yellow }
        default { Write-Host $entry -ForegroundColor Green }
    }
}

# ============================
# RETRY
# ============================

function Invoke-PCXWithRetry {
    param([scriptblock]$ScriptBlock)

    for ($i = 1; $i -le $script:Config.RetryCount; $i++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($i -eq $script:Config.RetryCount) {
                throw $_
            }
            Start-Sleep $script:Config.RetryDelay
        }
    }
}

# ============================
# FILE VALIDATION (FIXED)
# ============================
function Test-PCXPackagePath {
    param([string]$Path)

    $cleanPath = $Path.Trim()

    try {

        # Validate directory using pure .NET
        if (-not [System.IO.Directory]::Exists($cleanPath)) {
            throw "Path not accessible: $cleanPath"
        }

        # Enumerate files safely
        $items = [System.IO.Directory]::GetFiles($cleanPath) | ForEach-Object {
            [System.IO.FileInfo]$_
        }

        if (-not $items -or $items.Count -eq 0) {
            throw "No files found in package path: $cleanPath"
        }

        return $items
    }
    catch {
        throw "File enumeration failed on path: $cleanPath | $($_.Exception.Message)"
    }
}

# ============================
# METADATA
# ============================

function Get-PCXPackageMetadata {
    param([string]$Path)

    $clean = $Path.TrimEnd("\")
    $parts = $clean -split "\\"

    $company = $parts[-3]
    $raw = $parts[-1]

    $versionMatch = [regex]::Match($raw, '\d+(\.\d+)+')
    $version = if ($versionMatch.Success) { $versionMatch.Value } else { "1.0" }

    $product = $raw -replace [regex]::Escape($version), ""
    $product = $product -replace '[\.\-_]', ' '
    $product = ($product -replace '\s+', ' ').Trim()
    $product = $product -replace [regex]::Escape($company), ""
    $product = ($product -replace '\s+', ' ').Trim()

    $name = "$company $product $version"
    $packagename = "PKG $name"

    return @{
        Name        = $name
        PackageName = $packagename
        Company     = $company
        Product     = $product
        Version     = $version
    }
}

# ============================
# INSTALLER
# ============================

function Get-PCXInstaller {
    param($Files)

    $msi = $Files | Where-Object Extension -eq ".msi" | Select-Object -First 1
    if ($msi) { return $msi }

    $exe = $Files | Where-Object Extension -eq ".exe" | Select-Object -First 1
    if ($exe) { return $exe }

    throw "No installer found"
}

# ============================
# COMMAND BUILDER (FIXED SAFE FILE CHECK)
# ============================

function Get-PCXCommandLine {
    param(
        [string]$Path,
        [string]$Type,
        $Installer
    )

    $map = @{}

    # Safe filesystem enumeration (avoids SCCM PSDrive/provider issues)
    [System.IO.Directory]::GetFiles($Path) | ForEach-Object {

        $file = [System.IO.FileInfo]$_

        $map[$file.Name.ToLower()] = $file
    }

    switch ($Type) {

        "Install" {

            if ($map.ContainsKey("install.bat")) {
                return "cmd.exe /c install.bat"
            }

            if ($Installer.Extension -eq ".msi") {
                return "msiexec /i `"$($Installer.Name)`" /qn"
            }

            return "$($Installer.Name) /S"
        }

        "Uninstall" {

            if ($map.ContainsKey("uninstall.bat")) {
                return "cmd.exe /c uninstall.bat"
            }

            if ($Installer.Extension -eq ".msi") {
                return "msiexec /x `"$($Installer.Name)`" /qn"
            }

            return "$($Installer.Name) /uninstall /S"
        }

        "Upgrade" {

            if ($map.ContainsKey("upgrade.bat")) {
                return "cmd.exe /c upgrade.bat"
            }

            return $null
        }

        "OSD" {

            if ($Installer.Extension -eq ".msi") {
                return "msiexec /i `"$($Installer.Name)`" /qn"
            }

            return "$($Installer.Name)"
        }
    }
}

# ============================
# PROGRAM NAME FORMAT
# ============================
function Get-ProgramNames {
    param(
        [Parameter(Mandatory)]
        [string]$PackageName
    )

    [PSCustomObject]@{
        Available = "$PackageName [AVAILABLE]"
        Install   = "$PackageName [INSTALL]"
        Uninstall = "$PackageName [UNINSTALL]"
    }
}

function Get-CollectionNames {
    param(
        [Parameter(Mandatory)]
        [string]$PackageName
    )

    return [PSCustomObject]@{
        Available = "$PackageName [AVAILABLE]"
        Install   = "$PackageName [INSTALL]"
        Uninstall = "$PackageName [UNINSTALL]"
        Exception = "$PackageName [EXCEPTION]"
    }
}

# ============================
# UPGRADE CHECK
# ============================

function Test-PCXHasUpgrade {
    param([string]$Path)

    Test-Path (Join-Path $Path "upgrade.bat")
}

# ============================
# SCCM (DO NOT CHANGE - AS REQUESTED)
# ============================

function Get-PCXCMSiteCode {
    (Get-WmiObject -Namespace root\SMS -Class SMS_ProviderLocation).SiteCode | Select-Object -First 1
}

function Get-PCXCMProviderMachineName {
    (Get-WmiObject -Namespace root\SMS -Class SMS_ProviderLocation |
    Where-Object ProviderForLocalSite -eq $true).Machine
}

function Connect-PCXCMSite {
    param (
        [string]$SiteCode = $(Get-PCXCMSiteCode),
        [string]$ProviderMachineName = $(Get-PCXCMProviderMachineName)
    )

    $initParams = @{}

    if ((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams
    }

    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    Set-Location "$($SiteCode):\" @initParams
}

function New-PCXCMFolder {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$Path,

        [Parameter(Mandatory = $false, Position = 1)]
        [string]$Name
    )

    begin {
        Write-Verbose "********** BEGIN BLOCK **********"
    }

    process {
        Write-Verbose "********** Function Begin **********"

        try {
            # -------------------------------
            # Step 1: Detect and extract SiteCode
            # -------------------------------
            $siteCode = $null
            $cleanPath = $null

            if ($Path -match '^[A-Za-z0-9]{3}:\\') {
                # Path includes PSDrive (e.g., PS1:\...)
                $siteCode = $Path.Substring(0, 3)
                $cleanPath = $Path.Substring(4)
                Write-Verbose "Detected PSDrive in path: $siteCode"
            }
            else {
                # No PSDrive → use function
                $siteCode = Get-PCXCMSiteCode
                if (-not $siteCode) {
                    throw "Failed to retrieve SCCM Site Code."
                }
                $cleanPath = $Path
                Write-Verbose "Using detected SiteCode: $siteCode"
            }

            # -------------------------------
            # Step 2: Ensure ConfigMgr Module + PSDrive
            # -------------------------------
            if (-not (Get-PSDrive -Name $siteCode -ErrorAction SilentlyContinue)) {

                Write-Verbose "PSDrive '$siteCode' not found. Attempting to initialize..."

                $cmModulePath = Join-Path $ENV:SMS_ADMIN_UI_PATH "..\ConfigurationManager.psd1"

                if (-not (Test-Path $cmModulePath)) {
                    throw "ConfigurationManager module not found. Install SCCM Console."
                }

                Import-Module $cmModulePath -ErrorAction Stop
                Write-Verbose "ConfigurationManager module loaded."

                try {
                    Set-Location "$siteCode`:" -ErrorAction Stop
                    Write-Verbose "Connected to site drive: $siteCode"
                }
                catch {
                    throw "Failed to switch to PSDrive '$siteCode'. Verify site code."
                }
            }

            $rootPath = "$siteCode`:"
            
            # -------------------------------
            # Step 3: Normalize Path
            # -------------------------------
            $cleanPath = $cleanPath.Trim('\')

            if ([string]::IsNullOrWhiteSpace($cleanPath)) {
                throw "Path cannot be empty."
            }

            $segments = ($cleanPath -split '\\') | Where-Object { $_ }

            Write-Verbose "Normalized Path: $cleanPath"
            Write-Verbose "Segments: $($segments -join ' -> ')"

            # -------------------------------
            # Step 4: Create Path Step-by-Step
            # -------------------------------
            $currentPath = $rootPath

            foreach ($folder in $segments) {
                $nextPath = Join-Path $currentPath $folder

                if (-not (Test-Path $nextPath)) {
                    if ($PSCmdlet.ShouldProcess($nextPath, "Create folder")) {
                        New-Item -Path $currentPath -Name $folder -ItemType Directory -ErrorAction Stop
                        Write-Verbose "Created: $nextPath"
                    }
                }
                else {
                    Write-Verbose "Exists: $nextPath"
                }

                $currentPath = $nextPath
            }

            # -------------------------------
            # Step 5: Handle Optional Name
            # -------------------------------
            if ($Name) {
                if ([string]::IsNullOrWhiteSpace($Name)) {
                    throw "Folder name cannot be empty."
                }

                $finalPath = Join-Path $currentPath $Name

                if (-not (Test-Path $finalPath)) {
                    if ($PSCmdlet.ShouldProcess($finalPath, "Create folder")) {
                        New-Item -Path $currentPath -Name $Name -ItemType Directory -ErrorAction Stop
                        Write-Verbose "Created final folder: $finalPath"
                    }
                }
                else {
                    Write-Verbose "Final folder already exists: $finalPath"
                }
            }
            else {
                # No Name → full path already created
                $finalPath = $currentPath
                Write-Verbose "No child name provided. Full path ensured."
            }

            # -------------------------------
            # Step 6: Return Result
            # -------------------------------
            return [PSCustomObject]@{
                Success  = $true
                Path     = $finalPath
                SiteCode = $siteCode
            }
        }
        catch {
            Write-Error "Failed: $($_.Exception.Message)"

            return [PSCustomObject]@{
                Success = $false
                Error   = $_.Exception.Message
            }
        }
    }

    end {
        Write-Verbose "********** END BLOCK **********"
    }
}

# ============================
# ADD PROGRAM
# ============================

function Add-PCXProgram {
    param(
        [string]$PackageName,
        [string]$Type,
        [string]$CommandLine,
        $Platforms
    )

    $name = "$PackageName [$Type]"

    # Default values
    $runType = "WhetherOrNotUserIsLoggedOn"
    $userInteraction = $false
    $runMode = "RunWithAdministrativeRights"

    # Special handling for AVAILABLE
    if ($Type -eq "Available") {
        $runType = "OnlyWhenUserIsLoggedOn"
        $userInteraction = $true
    }

    # Create Program
    Invoke-PCXWithRetry {
        New-CMProgram `
            -PackageName $PackageName `
            -StandardProgramName $name `
            -CommandLine $CommandLine `
            -AddSupportedOperatingSystemPlatform $Platforms `
            -RunMode $runMode `
            -ProgramRunType $runType `
            -UserInteraction $userInteraction `
            -RunType Normal `
            -DiskSpaceRequirement 5 `
            -DiskSpaceUnit GB `
            -Duration 20
    }

    # Post config ONLY for Available
    if ($Type -eq "Available") {
        Invoke-PCXWithRetry {
            Set-CMProgram `
                -PackageName $PackageName `
                -ProgramName $name `
                -StandardProgram `
                -SuppressProgramNotification $false
        }
    }

    Write-PCXLog "$Type program created: $name"
}

function New-SCCMCollections {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Collections,

        [Parameter(Mandatory)]
        [string]$LimitingCollectionName
    )

    New-CMDeviceCollection -Name $Collections.Available -LimitingCollectionName $LimitingCollectionName
    New-CMDeviceCollection -Name $Collections.Install   -LimitingCollectionName $LimitingCollectionName
    New-CMDeviceCollection -Name $Collections.Uninstall -LimitingCollectionName $LimitingCollectionName
    New-CMDeviceCollection -Name $Collections.Exception -LimitingCollectionName $LimitingCollectionName

    Write-PCXLog "Collections created"
}

function Start-SCCMContentDistribution {
    param(
        [Parameter(Mandatory)]
        [string]$PackageName,

        [Parameter(Mandatory = $false)]
        [string]$DistributionPointGroupName = "All Mangalore DPs"
    )

    Start-CMContentDistribution `
        -PackageName $PackageName `
        -DistributionPointGroupName $DistributionPointGroupName

    Write-PCXLog "Content distributed to DP Group: $DistributionPointGroupName"
}

function New-SCCMDeployments {
    param(
        [Parameter(Mandatory)]
        [string]$PackageName,

        [Parameter(Mandatory)]
        [pscustomobject]$Programs,

        [Parameter(Mandatory)]
        [pscustomobject]$Collections,

        $DeadlineTime
    )

    $programComment = "$PackageName Program"

    New-CMPackageDeployment `
        -StandardProgram `
        -PackageName $PackageName `
        -CollectionName $Collections.Available `
        -Comment $programComment `
        -DeployPurpose Available `
        -ProgramName $Programs.Available

    if (-not $DeadlineTime) {
        $DeadlineTime = (Get-Date -Hour 20 -Minute 0 -Second 0).AddDays(30)
    }

    $schedule = New-CMSchedule -Start $DeadlineTime -Nonrecurring

    New-CMPackageDeployment `
        -StandardProgram `
        -PackageName $PackageName `
        -ProgramName $Programs.Install `
        -DeployPurpose Required `
        -CollectionName $Collections.Install `
        -Schedule $schedule

    New-CMPackageDeployment `
        -StandardProgram `
        -PackageName $PackageName `
        -ProgramName $Programs.Uninstall `
        -DeployPurpose Required `
        -CollectionName $Collections.Uninstall `
        -Schedule $schedule

    Write-PCXLog "Deployments created"
}

function Move-SCCMCollectionsToFolder {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Collections,

        [Parameter(Mandatory)]
        [pscustomobject]$meta
    )

    $siteCode = Get-PCXCMSiteCode

    $folder = "\DeviceCollection\Mphasis Application Deployment\$($meta.Company)\$($meta.Product)\$($meta.PackageName)"
    $folderPath = "$siteCode`:$folder"

    New-PCXCMFolder -Path $folder

    $collectionList = @(
        $Collections.Available,
        $Collections.Install,
        $Collections.Uninstall,
        $Collections.Exception
    )

    foreach ($collection in $collectionList) {

        $collectionObject = Get-CMDeviceCollection -Name $collection

        Move-CMObject `
            -FolderPath $folderPath `
            -InputObject $collectionObject

        Write-PCXLog "Moved Collection: $collection"
    }
}

function Move-SCCMPackageToFolder {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$meta
    )

    $siteCode = Get-PCXCMSiteCode

    $folder = "\Package\Application Installation\$($meta.Company)\$($meta.Product)"
    
    $folderPath = "$siteCode`:$folder"

    New-PCXCMFolder -Path $folder

    $packageObject = Get-CMPackage -Name $meta.PackageName -Fast

    Move-CMObject `
        -FolderPath $folderPath `
        -InputObject $packageObject

    Write-PCXLog "Moved Package: $($meta.PackageName)"
}

function Set-SCCMCollectionRules {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Collections
    )

    Add-CMDeviceCollectionIncludeMembershipRule `
        -CollectionName $Collections.Exception `
        -IncludeCollectionName $Collections.Uninstall

    Add-CMDeviceCollectionExcludeMembershipRule `
        -CollectionName $Collections.Install `
        -ExcludeCollectionName $Collections.Exception

    Write-PCXLog "Collection membership rules configured"
}

# ============================
# MAIN
# ============================

function Create-PCXPackage {

    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Language = "EN-US",
        [string]$DPGroup = "All Mangalore Dps",
        [string]$LimitingCollectionName = "ALL Systems"
    )

    try {
        Clear-Host
        Write-PCXLog "===== START ====="

        $files = Test-PCXPackagePath $Path
        $installer = Get-PCXInstaller $files
        $meta = Get-PCXPackageMetadata $Path

        $programs = Get-ProgramNames -PackageName $meta.PackageName
        $collections = Get-CollectionNames -PackageName $meta.PackageName

        Write-PCXLog "Package: $($meta.PackageName)"
        Write-PCXLog "Installer: $($installer.Name)"

        Connect-PCXCMSite

        $platforms = Get-CMSupportedPlatform -Fast | Where-Object {
            $_.DisplayText -like "*Windows 11*"
        }

        Invoke-PCXWithRetry {
            New-CMPackage -Name $meta.PackageName -Manufacturer $meta.Company -Version $meta.Version -Language $Language -Path $Path
        }

        Write-PCXLog "Package created"

        # INSTALL
        Add-PCXProgram $meta.PackageName "Install" (Get-PCXCommandLine $Path "Install" $installer) $platforms

        # AVAILABLE (NEW)
        Add-PCXProgram `
            -PackageName $meta.PackageName `
            -Type "Available" `
            -CommandLine (Get-PCXCommandLine $Path "Install" $installer) `
            -Platforms $platforms

        # UNINSTALL
        Add-PCXProgram $meta.PackageName "Uninstall" (Get-PCXCommandLine $Path "Uninstall" $installer) $platforms

        # UPGRADE (optional)
        if (Test-PCXHasUpgrade $Path) {
            $upCmd = Get-PCXCommandLine $Path "Upgrade" $installer
            if ($upCmd) {
                Add-PCXProgram $meta.PackageName "Upgrade" $upCmd $platforms
            }
        }

        # OSD
        Add-PCXProgram $meta.PackageName "OSD" (Get-PCXCommandLine $Path "OSD" $installer) $platforms

        Invoke-PCXWithRetry {
            Start-CMContentDistribution -PackageName $meta.PackageName -DistributionPointGroupName $DPGroup
        }

        New-SCCMCollections `
            -Collections $collections `
            -LimitingCollectionName $LimitingCollectionName

        $DeadlineTime = (Get-Date -Hour 20 -Minute 0 -Second 0).AddDays(30)

        New-SCCMDeployments `
            -PackageName $meta.PackageName `
            -Programs $programs `
            -Collections $collections `
            -DeadlineTime $DeadlineTime

        
        Set-SCCMCollectionRules `
            -Collections $collections
        
        Move-SCCMCollectionsToFolder `
            -Collections $collections `
            -meta $meta

        Move-SCCMPackageToFolder `
            -meta $meta
        
        Write-PCXLog "SUCCESS: $($meta.PackageName)"
    }
    catch {
        Write-PCXLog "FAILED: $($_.Exception.Message)" "ERROR"
        throw
    }
    finally {
        Write-PCXLog "Create Pacakge execution completed"
    }
}

# ============================
# EXECUTION
# ============================

Create-PCXPackage -Path "\\192.168.25.214\Package_source\Applications\Igor Pavlov\7zip\7zip 26.0.0\"
#Create-PCXPackage -Path "\\192.168.25.214\Package_source\Applications\Igor Pavlov\7zip\7zip 26.0.1\"