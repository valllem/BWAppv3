# Load external assemblies
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#################
# Launch Checks #
#################

######################
# App Location Check #
######################
$Path = Get-Location
$bin = "$Path" + "\" + "bin"
$Pathfile = $bin + "\path.bwapp"
$temp = "C:\Temp"
$tempfile = "C:\Temp\path.bwapp"
if (Test-Path -Path $tempfile) 
    {
    Set-Content -path "$tempfile" "$Path"
    }
else 
{
    New-Item -ItemType Directory -Force -Path $temp
    Set-Content -path "$tempfile" "$Path"
}


##################
# Settings Check #
##################
write-host 'Checking Configuration...'
if (Test-Path -Path ".\bin\config.bwapp") 
    {
    write-host -ForegroundColor Green 'OK'
    Get-Content -Path ".\bin\config.bwapp" | Select-String -SimpleMatch console
    }
else 
{
    write-host -ForegroundColor yellow 'Configuring'
    New-Item -ItemType Directory -Force -Path ".\bin"
    set-content ".\bin\config.bwapp" "console = enabled
version = 3
Build = '2003.28'
Path = '$Path'"
    
}

#[void][System.Windows.Forms.MessageBox]::Show("Update Available", "BWApp Launcher")
powershell.exe .\BWApp.ps1

