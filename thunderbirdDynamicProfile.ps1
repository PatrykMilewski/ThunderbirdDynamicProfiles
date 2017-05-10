<#

.SYNOPSIS
This script functionallity is to make thunderbird's mail profile, follow user in domain. It can import or export profile to/from user's home folder. Please notice that, it needs also some GP objects to make whole functionallity run correctly.

.PARAMETER mode
It's needed to choose, if script has to import or export user's profile.

.PARAMETER homeFolderLetter
Disc letter, that home folder is mapped to. By default it is taken from environemtal variable $env:homeDrive.

.PARAMETER homeFolder
Path to home folder joined with subfolder "Thunderbird", for example "H:\Thunderbird". Profile files will be saved there.

.PARAMETER logFileName
Name of log file.

.PARAMETER logFileLocation
Localisation of log file.

.PARAMETER localFolder
Localisation of Thunderbird's users data. It's folder, where we get our profile from.

.PARAMETER extensionsZipName
Name of compressed archive, where whole folder extensions from user's profile, will be stored.

.PARAMETER compressionProgram
You can select the program, that will be used for compression. Option "builtIn" means default PS compression cmdlet "Compress-Archive". By default script is trying to find the best one, where Compres/Expand-Archive > 7zipLocal. You can also add your custom compression program, but then you need to do some changes in script's code. Check compressFiles, expandFiles and findCompressionProgram functions.

.PARAMETER 7zipLocalisation
Non-default localistion of 7z.exe.

.PARAMETER emailNotifyFrom
Name of email address from which notification will be send.

.PARAMETER emailNotifyTo
Email of notification target.

.PARAMETER emailNotifyPassword
Password of email address to mail emailNotifyFrom.

.PARAMETER emailSmtpServer
Address of smtp server, for sending notifications.

.PARAMETER emailSmtpPort
Port of smtp server, for sending notifications.

.PARAMETER dontCopyExtensions
If you call this parameter, extensions will be no imported/exported.

.PARAMETER dontSendNotifications
If you call this parameter, notifications through email will be not send.

.PARAMETER taskTimeoutWarning
If script runs more than this parameter value (in seconds), notfication email will be send.

.PARAMETER extensionsMaxSize
Max size in bytes of extensions folder, which we can found in user's thunderbird profile. To make script more fail safe, script won't copy extensions, if they are bigger than max size.


.OUTPUTS
This script only returns error codes.
Error codes are organised by binary number, if there is error code number 1, then the youngest bit of return code will be "1", for example:
000001 - error code number 1,
010001 - error codes number 1 and 5,
000000 - no errors.
Error codes list:
1 - home folder drive is unreachable.
2 - thunderbird's application data with user's profile files is unreachable.
3 - there is no compression programs/cmdlets on this computer, extensions couldn't be imported/exported.
4 - profile.ini file is empty.
5 - profile.ini doesn't contain profile name, it may be broken.
6 - build of Lightening extensions was not found in file apps.ini, and it may be broken, or Lightening is not installed.
7 - build of Lightening extensions was not found in file lighteningBuild.txt, or file doesn't exist.
8 - creation of file's array throwed exception (not extensions files), when script was trying to import profile.
9 - expanding of extensions archive throwed exception, when script was trying to import profile.
10 - removing extensions archive throwed exception, when script was trying to import profile.
11 - exception was throwed, when script was trying to compress extensions, or create array of extensions files to export.
12 - extensions folder was too big, and script aborted exporting extensions.
13 - exception was throwed, when script tried to import profiles.ini file.
14 - exception was throwed, when script tried to export profiles.ini file.
15 - mode of script was uncorrect, only availiable options are "import" and "export".
16 - problem with sending email notification
17 - home drive is the same as system drive, it means, there is no home drive

#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True, Position=1)]
	 [String]$mode,
	[Parameter(Mandatory=$False)]
	 [String]$homeFolderLetter = $env:homeDrive,
	 [String]$homeFolder = (Join-Path $homeFolderLetter "Thunderbird"),
	 [String]$logFileName = "thunderbirdProfilesLogs.txt",
	 [String]$logFileLocation = (Join-Path $homeFolderLetter "Logs\$logFileName"),
	 [String]$localFolder = "$env:APPDATA\Thunderbird",
	 [String]$extensionsZipName = "extensions.zip",
	[ValidateSet('7zipLocal', 'builtIn', 'yourCmdlet')]
	 [String]$compressionProgram,
	 [String]$emailNotifyFrom = "powiadomienia.madler@gmail.com",
	 [String]$emailNotifyTo = "andrzej.milewski@gmail.com",
	 [String]$emailNotifyPassword,
	 [String]$emailSmtpServer = "smtp.gmail.com",
	 [String]$emailSmtpPort = "587",
	 [Switch]$dontCopyExtensions = $False,
	 [Switch]$dontSendNotifications,
	 [Int]$taskTimeoutWarning = 60,
	 [Int]$extensionsMaxSize = 50000000
)

Set-Variable -Name exitCode -Value 0 -Scope script

function addLog($newLog) {
	"[" + (Get-Date) + "] " + $newLog | Add-Content $logFileLocation
}

function exitScript($executionTime) {
	if ($executionTime -ne $null) {
		if ($Script:exitCode -ne 0) {
			$log = "Script exited with error code: $Script:exitCode after " + $executionTime.TotalSeconds + " seconds of execution in $mode mode."
		}
		else {
			$log = "Script exited after " + $executionTime.TotalSeconds + " seconds of execution in $mode mode."
		}
	}
	else {
		if ($Script:exitCode -ne 0) {
			$log = "Script exited with error code: $Script:exitCode in $mode mode."
		}
		else {
			$log = "Script exited in $mode mode."
		}
	}
	addLog $log
	exit $Script:exitCode
}

function findCompressionProgram() {
	if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) {
		$Script:compressionProgram = 'builtIn'
	}
	elseif (Test-Path ${env:ProgramFiles(x86)}\7-zip\7z.exe) {
		$7zipLocalisation = ${env:ProgramFiles(x86)} + "\7-zip\7z.exe"
		$Script:compressionProgram = '7zipLocal'
	}
	elseif (Test-Path $env:ProgramFiles\7-zip\7z.exe) {
		$7zipLocalisation = $env:ProgramFiles + "\7-zip\7z.exe"
		$Script:compressionProgram = '7zipLocal'
	}
	elseif ($compressionProgram -and (Get-Command $compressionProgram -errorAction SilentlyContinue)) {
		return
	}
	else {
		addLog "No archive program/cmdlet found. Extensions couldn't be moved."
		$dontCopyExtensions = $True
		$Script:exitCode += 4
	}
}

function compressFiles($source, $destination) {
	switch ($Script:compressionProgram) {
		'7zipLocal' {
			try {
				$arguments = "a -bd -aoa -mx1 $destination $source"
				$returnValue = Start-Process -FilePath $7zipLocalisation -ArgumentList $arguments -WindowStyle Hidden -ErrorAction Stop -Wait
				if (($returnValue.ExitCode -ne 0) -and ($returnValue.ExitCode -ne $null)){
					$exitErrorCode = $returnValue.ExitCode
					throw "7zip returned exit error code: $exitErrorCode."
				}
			}
			catch {
				throw $Error[0]
			}
		}
		'builtIn' {
			try {
				Compress-Archive -Path $source -DestinationPath $destination -CompressionLevel Fastest -Force -ErrorAction Stop
			}
			catch {
				throw $Error[0]
			}
		}
		default {
			throw "No compression cmdlet found. Please edit switch block in function compressFiles code, if you wish to add your custom archive compression cmdlet."
		}
	}
}

function expandFiles($source, $destination) {
	switch ($compressionProgram) {
		'7zipLocal' {
			try {
				$arguments = "x -bd -aoa $source -o$destination"
				$returnValue = Start-Process -FilePath $7zipLocalisation -ArgumentList $arguments -WindowStyle Hidden -ErrorAction Stop -Wait
				if (($returnValue.ExitCode -ne 0) -and ($returnValue.ExitCode -ne $null)){
					$exitErrorCode = $returnValue.ExitCode
					throw "7zip returned exit error code: $exitErrorCode."
				}
			}
			catch {
				throw $Error[0]
			}
		}
		'builtIn' {
			try {
				Expand-Archive -Path $source -DestinationPath $destination -ErrorAction Stop
			}
			catch {
				throw $Error[0]
			}
		}
		default {
			throw "No compression cmdlet found. Please edit switch block in function expandFiles code, if you wish to add your custom archive expanding cmdlet."
		}
	}
}

function getProfileName($source) {
	$profileName = "NULL"
	try {
		$profilesIni = Get-Content $source\profiles.ini -ErrorAction Stop
		if ($profilesIni) {
			$split = $profilesIni[6].Split("{/}")
			$profileName = $split[1]
		}
		else {
			addLog "The file prfoiles.ini is empty. Couldn't move Thunderbird profile."
			$Script:exitCode += 8
		}
	}
	catch {
		$profileName = "NULL"
		addLog "Profile name not found inside the file profiles.ini. Couldn't move Thunderbird profile."
		$Script:exitCode += 16
	}
	return $profileName
}

function getLighteningBuild($extensionDir, $sourceDir) {
	$readExtensionDirs = [System.IO.File]::OpenText($extensionDir + "\extensions.ini")
	$foundTable = 0
	while($null -ne ($lineExtension = $readExtensionDirs.ReadLine())) {
		if ($foundTable -eq 1) {
			$lineExtension = $lineExtension.Split("{=}")
			$extensionPath = $sourceDir + "\{" + $lineExtension[2] + "}\app.ini"
			if (Test-Path $extensionPath) {
				$readAppVersion = [System.IO.File]::OpenText($extensionPath)
				while ($null -ne ($lineVersion = $readAppVersion.ReadLine())) {
					if ($lineVersion -Match "BuildID=") {
						$lineVersion = $lineVersion.Split("{=}")
						$readAppVersion.close()
						return $lineVersion[1]
					}
				}
			}
		}
		if ($lineExtension -eq "[ExtensionDirs]") {
			$foundTable = 1
		}
	}
	$readExtensionDirs.close()
	addLog "Build ID of extension Lightening not found in file app.ini."
	$Script:exitCode += 32
	return 0
}

function readLighteningBuild($source) {
	$build = 0
	if (Test-Path $source\lighteningBuild.txt) {
		$readLighteningBuild = [System.IO.File]::OpenText($source + "\lighteningBuild.txt")
		while ($null -ne ($lighteningBuild = $readLighteningBuild.ReadLine())) {
			$build = $lighteningBuild
			$readLighteningBuild.close()
			return $build
		}
		$readLighteningBuild.close()
	}
	else {
		return $build
	}

	if ($build -eq 0) {
		addLog "Lightening build id not found in file lighteningBuild.txt. It may be broken."
		$Script:exitCode += 64
	}
	return $build
}

function importExtensions($source, $destination, $profileName, $checkIfExist) {
	if (!(Test-Path $destination\extensions)) {
		New-Item $destination\extensions -ItemType Directory -Force
	}
	try {
		$filesList = New-Object System.Collections.ArrayList
		$filesList.AddRange(("extensions.json", "extensions.ini", "addons.json", "$extensionsZipName"))
		$i = 0
		ForEach ($file in $filesList) {
			if (Test-Path $source\$file) {
				if ($i -eq 3) {
					$directoryInfo = Get-ChildItem $destination\extensions | Measure-Object

					if ($checkIfExist -eq $True -and $directoryInfo.count -eq 0) {
						Copy-Item $source\$file $destination\extensions -Force
					}
					elseif ($checkIfExist -eq $False) {
						Copy-Item $source\$file $destination\extensions -Force
					}
				}
				else {
					if ($checkIfExist -eq $True -and !(Test-Path $destination\$file)) {
						Copy-Item $source\$file $destination -Force
					}
					elseif ($checkIfExist -eq $False) {
						Copy-Item $source\$file $destination -Force
					}
				}
			}
			else {
				$log = "Error in importing extensions. File $file not found."
				addLog $log
			}
			$i++
		}
	}
	catch {
		addLog $error[0]
		$Script:exitCode += 128
	}
	if (Test-Path $destination\extensions\$extensionsZipName) {
		try {
			expandFiles -source "$destination\extensions\$extensionsZipName" -destination "$destination\extensions" -ErrorAction Stop
		}
		catch {
			addLog $error[0]
			$Script:exitCode += 256
		}
		try {
			Remove-Item $destination\extensions\$extensionsZipName
		}
		catch {
			addLog $error[0]
			$Script:exitCode += 512
		}
	}
}

function exportExtensions($source, $destiny) {
	$extensionsDir = new-object -Com scripting.filesystemobject
	$extensionsSize = $extensionsDir.getfolder($source + "\extensions").size
	if ($extensionsSize -lt $extensionsMaxSize) {
		$sourceNew = $source + "\extensions\*"
			try {
			compressFiles -source $sourceNew -destination "$source\$extensionsZipName" -ErrorAction Stop
			$filesList = New-Object System.Collections.ArrayList
			$filesList.AddRange(("$extensionsZipName", "extensions.json", "extensions.ini", "addons.json"))
			ForEach ($file in $filesList) {
				if (Test-Path $source\$file) {
					Copy-Item $source\$file $destiny -Recurse -Force -ErrorAction Stop
				}
				else {
					$log = "Error in exporting extensions. File $file not found."
					addLog $error[0]
				}
			}
		}
		catch {
			addLog $error[0]
			$Script:exitCode += 1024
			return 1
		}
	}
	else {
		addLog "Size of files in extensions directory is too big ($extensionsSize bytes, max $extensionsMaxSize)."
		$Script:exitCode += 2048
		return 1
	}
	return 0
}

function importFiles($source, $destination, $profileName) {
	if (Test-Path $source\profiles.ini) {
		Copy-Item $source\profiles.ini $destination -Recurse -Force -ErrorAction Stop
	}
	else {
		$log = "Error in importing file. File profiles.ini not found."
		$Script:exitCode += 4096
	}
	if (!(Test-Path $destination\Profiles\$profileName\calendar-data)) {
		New-Item $destination\Profiles\$profileName\calendar-data -ItemType Directory -Force
	}
	$filesList = New-Object System.Collections.ArrayList
	$filesList.AddRange(("key3.db", "deleted.sqlite", "local.sqlite", "abook.mab", "logins.json", "prefs.js", "storage.sdb", "cert8.db", "history.mab", "impab.mab"))
	$i = 0
	ForEach ($file in $filesList) {
		if (Test-Path $source\$file) {
			# if $i equals 1 or 2, we need to add \calendar-data to path
			if (($i -eq 1) -or ($i -eq 2)) {
				Copy-Item $source\$file $destination\Profiles\$profileName\calendar-data -Force
			}
			else {
				Copy-Item $source\$file $destination\Profiles\$profileName -Force
			}
		}
		else {
			$log = "Error in importing files. File $file not found."
			addLog $log
		}
		$i++
	}
}

function exportFiles($source, $destiny, $profileName) {
	$filesList = New-Object System.Collections.ArrayList
	$filesList.AddRange(("prefs.js", "logins.json", "key3.db", "storage.sdb", "abook.mab", "calendar-data\deleted.sqlite", "calendar-data\local.sqlite", "cert8.db", "history.mab", "impab.mab"))
	if (Test-Path $source\profiles.ini) {
		Copy-Item $source\profiles.ini $destiny -Force
	}
	else {
		addLog = "Error in exporting files. File profiles.ini not found."
		$Script:exitCode += 8192
	}
	
	ForEach ($file in $filesList) {
		if (Test-Path $source\Profiles\$profileName\$file) {
			Copy-Item $source\Profiles\$profileName\$file $destiny -Force
		}
		else {
			$log = "Error in exporting files. File $file not found."
			addLog $log
		}
	}
}

function importProfile($homeFolder, $localFolder, $homeFolderLetter) {
	$profileName = getProfileName $homeFolder
	if ($profileName -eq "NULL") {
		return
	}

	if (!(Test-Path $localFolder\Profiles\$profileName)) {
		New-Item $localFolder\Profiles\$profileName -ItemType Directory -Force
	}

	importFiles $homeFolder $env:APPDATA\Thunderbird $profileName

	if ($dontCopyExtensions -eq $False) {
		if ((Test-Path $localFolder\Profiles\$profileName\extensions.ini) -and (Test-Path $homeFolder\extensions.ini)) {
			$buildLocal = getLighteningBuild $localFolder\Profiles\$profileName $localFolder\Profiles\$profileName\extensions
			$buildExtern = readLighteningBuild $homeFolder

			if ($buildLocal -lt $buildExtern) {
				importExtensions $homeFolder $localFolder\Profiles\$profileName $profileName $False
				addLog "Swapped extensions files. Old build: $buildLocal New build: $buildExtern"
				return
			}
		}	
		importExtensions $homeFolder $localFolder\Profiles\$profileName $profileName $True
	}
}

function exportProfile($homeFolder, $localFolder, $homeFolderLetter) {
	if (!(Test-Path $homeFolder)) {
		New-Item $homeFolder -ItemType Directory
	}
	if ((Test-Path -Path $localFolder) -and (Test-Path -Path $homeFolderLetter)) {
		$profileName = getProfileName $localFolder
		if ($profileName -ne "NULL") {
			if ($dontCopyExtensions -eq $False) {
				$buildLocal = getLighteningBuild $localFolder\Profiles\$profileName $localFolder\Profiles\$profileName\extensions
				$buildExtern = readLighteningBuild $homeFolder
				if ($buildLocal -gt $buildExtern) {
					$returnValue = exportExtensions $localFolder\Profiles\$profileName $homeFolder
					if ($returnValue -eq 0) {
						New-Item $homeFolder\lighteningBuild.txt -Type File -Force -Value $buildLocal
					}
				}
			}
			exportFiles $localFolder $homeFolder $profileName
		}
	}
}


if (!(Test-Path $logFileLocation)) {
	New-Item $logFileLocation -ItemType File -Force
	addLog "Log file created."
}

if (!(Test-Path -Path $homeFolderLetter)) {
    addLog "Couldn't access home folder directory ($homeFolderLetter). Script will now exit."
	$Script:exitCode += 1
    exitScript
}

if ($homeFolderLetter -eq $env:SystemDrive) {
	addLog "There is no home drive associated to this account."
	$Script:exitCode += 65536
	exitScript
}

if (($mode -eq "export") -and !(Test-Path -Path $localFolder)) {
	addLog "Thunderbird files in application data not found. Script will now exit."
	$Script:exitCode += 2
	exitScript
}

if (!$compressionProgram) {
	findCompressionProgram
}
elseif ($compressionProgram -eq "7zipLocal" -and !($7zipLocalisation)) {
	if (Test-Path ${env:ProgramFiles(x86)}\7-zip\7z.exe) {
		$7zipLocalisation = ${env:ProgramFiles(x86)} + "\7-zip\7z.exe"
	}
	elseif (Test-Path $env:ProgramFiles\7-zip\7z.exe) {
		$7zipLocalisation = $env:ProgramFiles + "\7-zip\7z.exe"
	}
	else {
		findCompressionProgram
	}
}

$executionTime = Measure-Command {
	if ($mode -eq "import") {
		importProfile $homeFolder $localFolder $homeFolderLetter
	}
	elseif ($mode -eq "export") {
		exportProfile $homeFolder $localFolder $homeFolderLetter
	}
	else {
		$Script:exitCode += 16384
	}
}

if ($executionTime.TotalSeconds -gt $taskTimeoutWarning) {
	if ($dontSendNotifications -eq $False) { 
		$password = ConvertTo-SecureString $emailNotifyPassword -AsPlainText -Force
		$emailCredentials = New-Object pscredential ($emailNotifyFrom, $password)
		$subject = "Domena: ""$env:USERDOMAIN"" User: ""$env:USERNAME"" Komputer: ""$env:COMPUTERNAME"" Thunderbird dynamic profile notification."
		$body = "Użytkownik ""$env:username"" miał problem z wykonaniem skryptu thunderbirdDynamicProfile w trybie $mode. "
		$body += "Skrypt wykonywał się " + $executionTime.seconds + " sekund. Parametry wykonania: `n"
		$body += "taskTimeoutWarning: $taskTimeoutWarning`nhomeFolderLetter: $homeFolderLetter`nlocalFolder: $localFolder`n"
		$body += "extensionsMaxSize: $extensionsMaxSize`nPowerShell Version: " + $PSVersionTable.PSVersion.major + "`nexitCode: $Script:exitCode`n"
		try {
			Send-MailMessage -From $emailNotifyFrom -To $emailNotifyTo -Credential $emailCredentials -Subject $subject -Body $body -SmtpServer $emailSmtpServer -Port $emailSmtpPort -UseSsl -Encoding UTF8
		}
		catch {
			addLog $Error[0]
			$Script:exitCode += 32768
		}
	}
	$log = "Execution time warning, script was running to long (" + $executionTime.TotalSeconds + ")."
	addLog $log
}

exitScript $executionTime