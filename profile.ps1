# Set the colours for warnings/info as the script runs
$infoColour = "Yellow";
$errorColour = "Red";

function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

function IsWindowsTerminal ($childProcess) {
    if (!$childProcess) {
        return $false
    } elseif ($childProcess.ProcessName -eq 'WindowsTerminal') {
        return $true
    } else {
        return IsWindowsTerminal -childProcess $childProcess.Parent
    }
}

#================================================================================================================================================================
#================================================================================================================================================================
#====
#====   Main Profile script starts here
#====
#================================================================================================================================================================
#================================================================================================================================================================
Write-Host "$(Get-TimeStamp) Starting PoSh profile script" -ForegroundColor Green;


# Add python 3 to the path if it is installed
if ( (Test-Path "C:\Program Files\Python 3.9") -and (!$env:Path.Contains("C:\Program Files\Python 3.9")) ){
    Write-Host "$(Get-TimeStamp) Adding Python 3 to the path."
    $env:Path+= ";C:\Program Files\Python 3.9"
}
# Add python 3 to the path if it is installed
elseif ( (Test-Path "$env:LOCALAPPDATA\Programs\Python\Python39") -and (!$env:Path.Contains("$env:LOCALAPPDATA\Programs\Python\Python39")) ){
    Write-Host "$(Get-TimeStamp) Adding Python 3 to the path."
    $env:Path+= ";$env:LOCALAPPDATA\Programs\Python\Python39"
}



Write-Host "$(Get-TimeStamp) Checking PowerShellGet version...";
# Check that powershellget is at a good enough version...
if ( (Get-InstalledModule powershellget -ErrorAction SilentlyContinue).Version.split('.')[0] -lt 2 ){
    Write-Host "$(Get-TimeStamp) Module powershellget must be Version 2 or higher to install required modules. Update it using an admin PoSh session using the following command:" -ForegroundColor $errorColour;
    Write-Host "        Find-Module powershellget | Install-Module" -ForegroundColor $infoColour;
    return;
}

# Make sure the user modules folder is added to the PSModulePath
$userModulesFolder = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules"
if (!( (Test-Path $userModulesFolder) -and ($env:PSModulePath.Contains($userModulesFolder)) )){
    Write-Host "$(Get-TimeStamp) User modules folder not in PSModulePath. Adding '$userModulesFolder'."
    $env:PSModulePath += ";$userModulesFolder"
}




# Check if we have the powerline font 'Fira Code Retina' installed, and if not install it, set windows teminal to use it, and start a new tab...
Write-Host "$(Get-TimeStamp) Checking Fira Code font is installed...";
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null;
$firaCodeFonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families | Where-Object -Property Name -Like -Value "*fira*";
if ($firaCodeFonts -eq $null) {
    try {
        Write-Host "$(Get-TimeStamp) Installing Fira Code Retina font..." -ForegroundColor $infoColour;

        $tempFolder = New-TemporaryFile | % { Remove-Item $_; New-Item -ItemType Directory -Path ($_  -replace "\.tmp$", "") } | Select-Object -ExpandProperty FullName;
        $tempFile = "$tempFolder\FiraCode.zip";

        # Turn off the progress bar to speed up the download. See here - https://stackoverflow.com/a/43477248
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri "https://github.com/tonsky/FiraCode/releases/download/5.2/Fira_Code_v5.2.zip" -OutFile $tempFile;
        $ProgressPreference = 'Continue'

        Expand-Archive -Path $tempFile -DestinationPath $tempFolder;
        $fontFile = "$tempFolder\ttf\FiraCode-Retina.ttf"

        # Install the font. Based on this - https://gist.github.com/anthonyeden/0088b07de8951403a643a8485af2709b
        $Destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
        $Destination.CopyHere($fontFile, 0x10)

        # If we are using windows terminal, change the settings file to use the font and start a new tab...
        # if ($env:WT_SESSION -ne $null) {
        #     Write-Host "$(Get-TimeStamp) Detected Windows Terminal is the shell host, configuring it to use the font..." -ForegroundColor $infoColour;
        #     #TODO: Update the settings file to use the font
        # }
    } catch {
        Write-Host "$(Get-TimeStamp) Installing Fira Code Retina font failed" -ForegroundColor $errorColour;
    } finally {
        if (Test-Path -Path $tempFolder) {
            Remove-Item -Path $tempFolder -Recurse -Force;
        }
    }
}


Write-Host "$(Get-TimeStamp) Importing posh-git module...";
# Poshgit install / import
if ( -not (Test-Path '~\src\posh-git') ){
    Write-Host "$(Get-TimeStamp) Cloning posh-git" -ForegroundColor $infoColour;
    Push-Location ~;
    git clone 'https://github.com/dahlbyk/posh-git.git' '.\src\posh-git';
    Pop-Location;
}
Import-Module '~\src\posh-git\src\posh-git.psd1'



# Other modules install / import
Write-Host "$(Get-TimeStamp) Importing modules...";

$requiredModules = @(
    @{ name = "oh-my-posh"; version = "2.0.496" } # Some versions have been triggering the Zscaler virus checker
   ,@{ name = "PSReadLine" }
   ,@{ name = "posh-sshell" }
)
foreach ($module in $requiredModules) {
    #if ( (Get-InstalledModule -Name $module -ErrorAction SilentlyContinue) -eq $null) {
    $installedModule = Get-InstalledModule -Name $module.name -ErrorAction SilentlyContinue;
    
    # If no version is specified and the module is not installed, install it
    if ($installedModule -eq $null -and $module.version -eq $null) {
            Write-Host "$(Get-TimeStamp) Installing module: $($module.name)" -ForegroundColor $infoColour;
            Install-Module -Name $module.name -Scope CurrentUser -Confirm:$False;# -Force;
    
    # If the module is nto the correct version, install or update it
    } elseif ($installedModule -eq $null -and $module.version -ne $installedModule.Version) {
        if ($installedModule -eq $null) {
            Write-Host "$(Get-TimeStamp) Installing module: $($module.name) ($($module.version)) " -ForegroundColor $infoColour;
            Install-Module -Name $module.name -Scope CurrentUser -RequiredVersion $module.version -Confirm:$False;# -Force;
        } else {
            Write-Host "$(Get-TimeStamp) Updating module: $($module.name)" -ForegroundColor $infoColour;
            Update-Module -Name $module.name -Scope CurrentUser -Confirm:$False;# -Force;
        }

    # Else log a message that it's already installed
    } else {
        Write-Host "$(Get-TimeStamp) Module already installed: $($module.name)" -ForegroundColor $infoColour;
    }
    #}
    Import-Module -Name $module.name;
}





# Bash style auto-complete (PSReadLine is installed by default on win 10!)
Set-PSReadlineKeyHandler -Key Tab -Function MenuComplete

# Poshgit settings
$GitPromptSettings.DefaultPromptPrefix.Text = '`n';

# Oh-my-posh settings
Set-PoshPrompt -Theme Paradox;

# Change a few windows settings

# # 1. Dark mode on...
# Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ -Name AppsUseLightTheme -Value 0
# Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ -Name SystemUsesLightTheme -Value 0

# # 1. Wallpaper
# $wallpaperPath = Get-ChildItem ~\Pictures\DesktopBackground | Get-Random | Select-Object -ExpandProperty Fullname
# Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop\' -Name WallPaper -Value $wallpaperPath





Write-Host "$(Get-TimeStamp) Setting up helper functions";

# Variable with my custom folders. We generate goto functions for each one, so we can either use the variable for the actual path, or use the function
# to get there fast
$folders = [PSCustomObject]@{
    src = "~\src"
    srcCoreDatabase = "~\src\EDW - Accelerator Core\EDW - Accelerator Core Database"
};
$folders.PSObject.Properties | ForEach-Object { Invoke-Expression "Function GoTo-$($_.Name.SubString(0, 1).ToUpper())$($_.Name.SubString(1)) { cd `"$($_.Value)`" }" }

# Starts the ssh agent and loads the keys in ~/.ssh, prompting for passwords as required
function Start-SshAgentWithKeys {
    # Clear keys before loading
    ssh-add -D;

    # Start the SSH agent and load all the keys. This will prompt for passwords...
    Start-SshAgent;
    Get-ChildItem ~\.ssh\ | 
        Where-Object { $_.Name -NotLike '*.pub' -and $_.Name -ine 'config' -and $_.Name -ine "known_hosts" } | 
        ForEach-Object { Add-SshKey $_.FullName};

    # List the keys now loaded
    ssh-add -l
}

# Function to stop the autolock kicking in so fast!
function Start-PreventSleep {
    param (
        [Parameter(Mandatory=$true)]
        [String]
        $Until
    )
    
    $shell = New-Object -ComObject Wscript.Shell;

    while ($true) {
        $shell.SendKeys('{NUMLOCK}')
        $shell.SendKeys('{NUMLOCK}')

        if ( (Get-Date) -gt (Get-Date ($Until -replace '\.', ':')) ){
            break;
        }
        else {
            Start-Sleep -Seconds 60
        }
    }
}





Write-Host "$(Get-TimeStamp) Finished PoSh profile script";
