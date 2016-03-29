function New-NAVService
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [String]$DatabaseName,
    [Parameter(Mandatory)]
    [String]$ServiceTier
    )
  Process
  {
    $Config = New-NAVServiceConfig -Database $DatabaseName -ServiceTier $ServiceTier
    Install-NAVService -ServiceInstanceName $DatabaseName -ServiceTier $ServiceTier -ConfigFile $Config
    Set-ServiceCredential
  }
}

function Install-NAVService
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [String]$ServiceInstanceName,
    [Parameter(Mandatory)]
    [String]$ServiceTier,
    [Parameter(Mandatory)]
    [String]$ConfigFile
    )
  Process
  {
    $NAVServer = Join-Path $ServiceTier "Microsoft.Dynamics.Nav.Server.exe"
    $binPath = "\`"$NAVServer\`" `$$ServiceInstanceName config \`"$ConfigFile\`""
    $DisplayName = "Microsoft Dynamics NAV Server[$ServiceInstanceName]"
    $Dependencies = "HTTP/NetTcpPortSharing"
    $Name = "MicrosoftDynamicsNavServer`$$ServiceInstanceName"
    $ArgumentsList = "create `"$Name`" DisplayName= `"$DisplayName`" depend= $Dependencies binPath=`"$binPath`""
    Start-Process -FilePath "sc.exe" -ArgumentList $ArgumentsList -Verb "RunAs" -WindowStyle Hidden
  }
}

function New-NAVServiceConfig
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [String]$Database,
    [Parameter(Mandatory)]
    [String]$ServiceTier,
    [Parameter()]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$ConfigLocation
    )
  Process
  {
    $NewConfigLocation = $ConfigLocation
    if(!$NewConfigLocation){
      $NewConfigLocation = Join-Path $ServiceTier "Instances"
    }
    if(!Test-Path $NewConfigLocation -PathType Container){
      New-Item -Path $NewConfigLocation -ItemType Directory | Out-Null
    }
    $GUID = [GUID]::NewGUID().Guid
    $TempPath = Join-Path $env:temp $GUID
    New-Item -Path $TempPath -ItemType Directory | Out-Null
    $SettingsFile = Join-Path $TempPath "CustomSettings.config"
    [xml]$xml = Get-Content $(Join-Path $ServiceTier "CustomSettings.config")
    ($xml.appSettings.add | ? {$_.key.equals("DatabaseName")}).value = $Database
    ($xml.appSettings.add | ? {$_.key.equals("ServerInstance")}).value = $Database
    $xml.save($($SettingsFile.FullName))
    Copy-Item -Path $(Join-Path $ServiceTier "Tenants.config") -Destination $TempPath
    Copy-Item -Path $(Join-Path $ServiceTier "Microsoft.Dynamics.Nav.Server.exe.config") -Destination $(Join-Path $TempPath "$ServiceTier.config") -PassThru
    $SettingsFolder = Move-Item -Path $TempPath -Destination $(Join-Path $ConfigLocation $Database) -PassThru
    Get-Item $(Join-Path $SettingsFolder "$ServiceTier.config")
  }
}

function Remove-NAVService
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [String]$ServiceInstanceName
    )
  Process
  {
    if(!$ServiceInstanceName.StartsWith("MicrosoftDynamicsNavServer")){
      $ServiceInstanceName = "MicrosoftDynamicsNavServer`$$ServiceInstanceName"
    }
    $ServiceInstanceName = $ServiceInstanceName.ToLower()
    $Services = gwmi win32_service | ? {$_.Name.ToLower.Equals($ServiceInstanceName)} | measure
    if ($Services.Count == 0){
      throw "Service not found"
    }
    $ArgumentsList = "delete `"$ServiceInstanceName`" DisplayName= `"$DisplayName`" depend= $Dependencies binPath=`"$binPath`""
    Start-Process -FilePath "sc.exe" -ArgumentList $ArgumentsList -Verb "RunAs" -WindowStyle Hidden
  }
}

function Get-NAVDatabases
{
  [CmdletBinding()]
  Param
  (
    [Parameter()]
    [String]$SQLServerInstance
    )
  Process
  {
    Push-Location
    if(!$SQLServerInstance)
    {
      $SQLServerInstance = "localhost"
    }
    $databases = Invoke-SQLCmd "SELECT name FROM master.dbo.sysdatabases" -ServerInstance $SQLServerInstance
    foreach($database in $databases)
    {
      $dbVersion = Invoke-SQLCmd "SELECT databaseversionno FROM [$($database.name)].[dbo].[`$ndo`$dbproperty]" -ServerInstance $SQLServerInstance -ErrorAction SilentlyContinue
      if ($dbVersion) {
        [PSCustomObject]@{
          DatabaseName = $database.name
          DBVersion = $dbVersion.databaseversionno
        }  
      }
    }
    Pop-Location
  }
}

function New-Shortcut
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]$TargetPath,
    [Parameter(Mandatory)]
    [ValidateScript({!(Test-Path $_ -PathType Leaf)})]
    [String]$DestinationFile,
    [Parameter()]
    [String]$Arguments,
    [Parameter()]
    [String]$IconLocation,
    [Parameter()]
    [String]$Description,
    [Parameter()]
    [String]$WorkingDirectory
    )
  Process
  {
    $WshShell = New-Object -ComObject WScript.Shell
    if(!$DestinationFile.EndsWith(".lnk")){
      $DestinationFile += ".lnk"
    }
    $Shortcut = $WshShell.CreateShortcut($DestinationFile)
    if($TargetPath){
      $Shortcut.TargetPath = $TargetPath
    }
    if($Arguments){
      $Shortcut.Arguments = $Arguments
    }
    if($IconLocation){
      $Shortcut.IconLocation = $IconLocation
    }
    if($Description){
      $Shortcut.Description = $Description
    }
    if($WorkingDirectory){
      $Shortcut.WorkingDirectory = $WorkingDirectory
    }
    $Shortcut.Save()
    Get-Item $DestinationFile
  }
}

function Set-NAVServerLicense
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [String]$ServiceInstance,
    [Parameter()]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]$LicenseFile = $script:LicenseFileLocation,
    [Parameter()]
    [bool]$RestartNavService = $true
    )
  Process
  {
    $license = [Byte[]]$(Get-Content -Path $LicenseFile -Encoding Byte)

    Import-NAVServerLicense -ServerInstance $ServiceInstance -LicenseData $license -Database NavDatabase | Out-Null
    if($RestartNavService)
    {
      Write-Verbose "Restarting $ServiceInstance"
      Set-NAVServerInstance -ServerInstance $ServiceInstance -Restart | Out-Null
    }
  }
}

function Set-NAVServerPortSharing
{
  [CmdletBinding()]
  Param
  (
    )
  Process
  {
    $services = (gwmi win32_service | ? {$_.Name.StartsWith("MicrosoftDynamicsNavServer")}).Name
    foreach($service in $services){
      $serviceinfo = sc.exe qc $service
      foreach($serviceinformation in $serviceinfo){
        if($serviceinformation.Contains("DEPENDENCIES")){
          $dependencies = $serviceinformation.SubString($serviceinformation.IndexOf(": ") + 1)
        }
      }
      if(!$dependencies.Contains("NetTcpPortSharing")){
        $dependencies = $dependencies.Trim() + '/NetTcpPortSharing'
        sc.exe config $service depend= $dependencies
      }
    }
  }
}

function Set-ServiceCredential 
{
  [CmdletBinding()]
  Param
  (
    )
  Process 
  {
    $credentials = Get-Credential -UserName $(whoami) -Message "UserName & Password for services"
    $services = gwmi win32_service | ? {$_.Name.StartsWith("MicrosoftDynamicsNavServer")}
    ForEach ($service in $services) {
      $service.StopService()
      $service.Change($null, $null, $null, $null, $null, $null, $(whoami), $credentials.GetNetworkCredential.Password, $null, $null, $null)
      $service.StartService()
    }
  }
}

function Expand-NAVFolder
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]$ZipFile,
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$DestinationPath,
    [switch]$Open
    )
  Process
  {
    $ZipFile = (Get-Item $ZipFile).FullName
    $DestinationPath = (Get-Item $DestinationPath).FullName
    $GUID = [GUID]::NewGUID().Guid
    $TempPath = Join-Path $env:temp $GUID
    if(-not (Test-Path -Path $TempPath -PathType Container)){
      New-Item -Path $TempPath -ItemType Directory | Out-Null
    }
    Expand-Archive -Path $ZipFile -DestinationPath $TempPath
    $NAVProductVersion = Get-ChildItem -Path $TempPath -Include 'finsql.exe' -Recurse | Select -First 1 | Select -ExpandProperty VersionInfo | Select -ExpandProperty ProductVersion
    $VersionPath = Join-Path $DestinationPath $NAVProductVersion
    if(-not (Test-Path -Path $VersionPath -PathType Container)){
      New-Item -Path $VersionPath -ItemType Directory | Out-Null
    }
    $NAVFolderVersion = Get-ChildItem -Attributes D -Name -Path $(Join-Path $TempPath 'RoleTailoredClient\program files\Microsoft Dynamics NAV\') | Select -First 1
    $NSTFolder = (New-Item -Path $(Join-Path $VersionPath 'NST') -ItemType Directory -Force).FullName
    $RTCFolder = (New-Item -Path $(Join-Path $VersionPath 'RTC') -ItemType Directory -Force).FullName
    Copy-Item -Path $(Join-Path $TempPath "RoleTailoredClient\program files\Microsoft Dynamics NAV\$NAVFolderVersion\RoleTailored Client\*") -Destination $RTCFolder -Recurse -Force
    Copy-Item -Path $(Join-Path $TempPath "ServiceTier\program files\Microsoft Dynamics NAV\$NAVFolderVersion\Service\*") -Destination $NSTFolder -Recurse -Force
    Copy-Item -Path $(Join-Path $TempPath "Installers\DK\RTC\PFiles\Microsoft Dynamics NAV\$NAVFolderVersion\RoleTailored Client\*") -Destination $RTCFolder -Recurse -Force
    Copy-Item -Path $(Join-Path $TempPath "Installers\DK\Server\PFiles\Microsoft Dynamics NAV\$NAVFolderVersion\Service\*") -Destination $NSTFolder -Recurse -Force
    Remove-Item -Path $TempPath -Recurse -Force
    if($Open){
      Invoke-Item $VersionPath
    }
  }
}

function New-NAVShellScript
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$RTCFolder,
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$NSTFolder,
    [Parameter(Mandatory)]
    [String]$ServiceInstanceName,
    [Parameter(Mandatory)]
    [String]$DestinationFile
    )
  Process 
  {
    $ModelTools = Join-Path $RTCFolder "Microsoft.Dynamics.Nav.Model.Tools.dll"
    $Management = Join-Path $NSTFolder "Microsoft.Dynamics.Nav.Management.dll"
    $scriptString = "Import-Module `"$($ModelTools.FullName)`"`n"
    $scriptString += "Import-Module `"$($Management.FullName)`"`n"
    $scriptString += "`$global:serviceInstance = `"$ServiceInstanceName`"`n"
    Out-File -FilePath $DestinationFile -Encoding "UTF8" $scriptString
    Get-File $DestinationFile
  }
}

function New-AdminShellScript
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]$ScriptFile,
    [Parameter(Mandatory)]
    [String]$DestinationFile
    )
  Process 
  {
    $scriptString = "Start-Process powershell -Verb RunAs -ArgumentList `"-NoExit`",`"-File $ScriptFile`""
    Out-File -FilePath $DestinationFile -Encoding "UTF8" $scriptString
    Get-File $DestinationFile
  }
}

function New-NAVDeveloperShortcut
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [String]$DeveloperExeFile,
    [Parameter(Mandatory)]
    [String]$ShotcutLocation,
    [Parameter()]
    [String]$id
    )
  Process
  {
    if($id){
      $Arguments = "id=`"$id`""
    } 
    $DeveloperExeFolder = Split-Path $DeveloperExeFile
    $IconFile = $DeveloperExeFile + ',0'
    New-Shortcut -TargetPath $DeveloperExeFile -DestinationFile $ShotcutLocation -Arguments $Arguments -IconLocation $IconFile -Description "Microsoft Dynamics NAV Developer Environment" -WorkingDirectory $DeveloperExeFolder
  }
}

function New-NAVRTCShortcut
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [String]$RTCExeFile,
    [Parameter(Mandatory)]
    [String]$ShotcutLocation,
    [Parameter(Mandatory)]
    [String]$ServiceInstance
    )
  Process
  {
    $Arguments = "`"DynamicsNAV://localhost:7046/$ServiceInstance//`""
    $RTCExeFolder = Split-Path $RTCExeFile
    $IconFile = $RTCExeFile + ',0'
    New-Shortcut -TargetPath $RTCExeFile -DestinationFile $ShotcutLocation -Arguments $Arguments -IconLocation $IconFile -Description "Microsoft Dynamics NAV Role Tailored Client" -WorkingDirectory $RTCExeFolder
  }
}

function New-PowershellShortcut 
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [String]$ScriptFile,
    [Parameter(Mandatory)]
    [String]$DestinationFile
    )
  Process 
  {
    $orig = Get-ChildItem -Path "$($env:userprofile)\AppData\Roaming\Microsoft\Windows\Start Menu" -Filter "Windows Powershell.lnk" -Recurse -Force | Select-Object -First 1
    Copy-Item $orig.FullName $DestinationFile
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($DestinationFile)
    $shortcut.TargetPath = "`"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`" -File `"$ScriptFile`""
    $shortcut.Save()
  }
}

function New-NAVShortcuts 
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$DestinationFolder,
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$RTCFolder,
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$ScriptFile,
    [Parameter(Mandatory)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [String]$ServiceInstanceName
    )
  Process 
  {
    $newFolder = New-Item $(Join-Path $DestinationFolder $ServiceInstanceName) -ItemType Directory -PassThru
    New-NAVDeveloperShortcut -DeveloperExeFile $(Join-Path $RTCFolder "finsql.exe") -ShotcutLocation $(Join-Path $newFolder "FinSQL") -id $ServiceInstanceName
    New-NAVDeveloperShortcut -DeveloperExeFile $(Join-Path $RTCFolder "finsql.exe") -ShotcutLocation $(Join-Path $newFolder "FinSQL - DEV") -id "$ServiceInstanceName-DEV"
    New-PowershellShortcut -ScriptFile $ScriptFile -DestinationFile $(Join-Path $newFolder "Shell")
    New-NAVRTCShortcut -RTCExeFile $(Join-Path $RTCFolder "Microsoft.Dynamics.Nav.Client.exe") -ShotcutLocation $(Join-Path $newFolder "RTC") -ServiceInstance $ServiceInstanceName
    Get-Item $newFolder
  }
}

Write-Host "$($MyInvocation.MyCommand) Loaded!"

Export-ModuleMember -Function * -Cmdlet *