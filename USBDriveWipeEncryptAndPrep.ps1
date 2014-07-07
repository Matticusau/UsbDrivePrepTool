#Requires -Version 4
#Requires -RunAsAdministrator
#Requires -Modules BitLocker

<#
    Formats, Encrypts and copies standard files to a USB drive for consultancy type work

    Written By: Matt Lavery
    Date:       07/05/2014


    Change History
    Version  Who          When           What
    --------------------------------------------------------------------------------------------------
    1.0      MLavery      07-May-2014    Initial Coding
    1.1      MLavery      06-Jun-2014    Improved progress bar and verbose logging
    1.2      MLavery      07-Jul-2014    Bug fixes and validates the bitlocker export path is protected

#>

<#
    .NOTES
        
        DISCLAIMER
        This Sample Code is provided for the purpose of illustration only and is not intended to be 
        used in a production environment.  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED 
        "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED 
        TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant 
        You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and 
        distribute the object code form of the Sample Code, provided that You agree: (i) to not use 
        Our name, logo, or trademarks to market Your software product in which the Sample Code is 
        embedded; (ii) to include a valid copyright notice on Your software product in which the 
        Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our 
        suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise 
        or result from the use or distribution of the Sample Code.
    
    .SYNOPSIS
        Formats, Encrypts and copies standard files to a USB drive for consultancy type work

    .DESCRIPTION
        Formats, Encrypts and copies standard files to a USB drive for consultancy type work

    .PARAMETER DriveLetter
        The drive letter of the removable USB drive to format, encrypt and then copy files to

    
    .EXAMPLE
        .\USBDriveWipeEncryptAndPrep.ps1

        Automatically detects any attached USBDrive and processes using that if accepted by the user

    .EXAMPLE
        "D:" | .\USBDriveWipeEncryptAndPrep.ps1

        Uses the value along the Pipeline for the Drive Letter of the USBDrive to process
    
    .EXAMPLE
        .\USBDriveWipeEncryptAndPrep.ps1 -DriveLetter "F:"

        Uses the value supplied to the DriveLetter parameter for the USBDrive to process

    .EXAMPLE
        .\USBDriveWipeEncryptAndPrep.ps1 -MountPoint "F:"

        Alternative to previous example as it uses the MountPoint parameter (alias) instead of the DriveLetter parameter name
    
    .OUTPUTS
        None
#>
[CmdletBinding(DefaultParameterSetName = "DriveLetter", SupportsPaging = $false, SupportsShouldProcess = $false, HelpURI = "http://blog.matticus.net")]
param(
    [Parameter(ParameterSetName="DriveLetter", Mandatory=$false, ValueFromPipeline=$true, Position=1, HelpMessage = "The drive letter of which drive to format.")]
    [Alias("MountPoint","dl")]
    [string]
    [ValidateNotNullOrEmpty()]
    $DriveLetter
    ,
    [Parameter(ParameterSetName="DriveLetter", Mandatory=$false, ValueFromPipeline=$false, Position=2, HelpMessage = "Use when you wish to build the configuration file.")]
    [switch]
    $Setup
)

begin 
{
    #Functions needed for the script
    function FormatElapsedTime ([Timespan]$TimeSpan)
    {
        Return [string]::Format("{0:00}:{1:00}.{2:00}", $TimeSpan.Hours, $TimeSpan.Minutes, $TimeSpan.Seconds);
    }
    function ValidateSettingsFile
    {
        if (!(Test-Path "$($Script:ScriptServerPath)\$($Script:ScriptName)_config.xml") -or $Setup -eq $true)
        {
            [xml]$Settings = "<Config ScriptName=`"$Script:ScriptName.ps1`" Version=`"1.0`"><Settings><CustomerToolsDir></CustomerToolsDir><BitLockerExportDir></BitLockerExportDir></Settings></Config>"
            Write-Host "`nWhoops.... " -ForegroundColor Red -BackgroundColor Black -NoNewline;
            Write-Host "Could not find configuration file or -Setup parameter supplied, you will be prompted for the configuration to create the file";
            Write-Host "`n`nConfig File Setup" -ForegroundColor Green;
            $Settings.Config.Settings.CustomerToolsDir = [string](Read-Host -Prompt "Enter a directory which contains files to copy to your USB drive [blank to skip this step]");
            $Settings.Config.Settings.BitLockerExportDir = [string](Read-Host -Prompt "Enter a directory to save the BitLocker2Go recovery key file to [required]");
            $Settings.Save("$($Script:ScriptServerPath)\$($Script:ScriptName)_config.xml");
            Write-Host "`nConfig saved.`n`n" -ForegroundColor Green;
            #flick the switch that we have performed setup to avoid looping
            #$Setup = $false;

            #revalidate the file
            #ValidateSettingsFile;
            Break;
        }
        else
        {
            [xml]$Settings = Get-Content "$($Script:ScriptServerPath)\$($Script:ScriptName)_config.xml";
            if ($Settings.Config.Settings.BitLockerExportDir.Trim().Length -le 0)
            {
                Write-Host "`nWhoops.... There is an error with the configuration script." -ForegroundColor Red -BackgroundColor Black;
                Write-Host "`nPlease either...";
                Write-Host "1) edit the configuration file $($Script:ScriptServerPath)\$($Script:ScriptName)_config.xml";
                Write-Host "2) run .\$($Script:ScriptName).ps1 -Setup";
                Write-Host "Then retry the script`n`n";
                Break;
            }
        }
    }
    function IsBitLockeredLocation([String]$Path)
    {
        Write-Debug "IsBitLockeredLocation $Path";
        #check that the path actually exists
        if ((Test-Path $Path) -eq $false){Return $false; Break;}

        Write-Debug "Getting bit vol object";
        #Check if this volume is protected with bitlocker 
        $BitLocVol = Get-BitLockerVolume -MountPoint "$((Get-Item $Path).Root)" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue;

        Return ($BitLocVol -ne $null);
    }

    #push the current location to the stack
    Push-Location;

    #get the settings
    $Script:ScriptName = $(Split-Path (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path -Leaf).Replace(".ps1","")
    $Script:ScriptServerPath = Split-Path (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path
    ValidateSettingsFile
    [xml]$Script:Settings = Get-Content "$($Script:ScriptServerPath)\$($Script:ScriptName)_config.xml"

    #$Script:CustomerTools = "C:\Users\mlavery\Documents\USBDrive\*"
    #$Script:BitLockerExportPath = "C:\Users\mlavery\SkyDrive @ Microsoft\Private Files\Bitlocker Recovery Keys\"
    $Script:CustomerTools = $Script:Settings.Config.Settings.CustomerToolsDir;
    $Script:BitLockerExportPath = $Script:Settings.Config.Settings.BitLockerExportDir;
    
    #Checking that the BitLocker key save path is protected
    if (IsBitLockeredLocation($Script:BitLockerExportPath))
    {
        Write-Debug "IsBitLockeredLocation returned true";
    }
    else
    {
        Write-Debug "IsBitLockeredLocation returned false";
        Write-Host "`nWhoops.... There is an error with the configuration script." -ForegroundColor Red -BackgroundColor Black;
        Write-Host "`nThe path specified as the BitLocker Key export path is either not valid or is not protected by BitLocker.";
        Write-Host "Storing the key on a non secured volume voids any protection gained through BitLocker.";
        Write-Host "`nPlease either...";
        Write-Host "1) edit the configuration file $($Script:ScriptServerPath)\$($Script:ScriptName)_config.xml";
        Write-Host "2) run .\$($Script:ScriptName).ps1 -Setup";
        Write-Host "Then retry the script`n`n";
        Break;
    }    
    
    #create a PSDrive
    Write-Verbose "Creating a PSDrive to reference the location of the Bitlocker export path";
    $Result = New-PSDrive -Name BitLockerKeys -PSProvider FileSystem -Root "$Script:BitLockerExportPath";

    #check that the drive was created successful
    if ($? -eq $false)
    {
        Write-Verbose "Error creating PSDrive";
        Write-Host "Could not create PSDrive for BitLockerKey export, script will exit" -ForegroundColor Red -BackgroundColor Black;
        Exit;
    }

    #Start the timer
    $Script:ElapsedTime = New-Object System.Diagnostics.Stopwatch;
    $Script:ElapsedTime.Start();

}

process
{
    $Script:Volume = @();

    #Get the details of the volume
    if ($Script:DriveLetter.Length -eq 0 -or $Script:DriveLetter -eq $null)
    {
        Write-Verbose "Getting the volume object for any removable drive";
        $Script:Volume = Get-Volume | Where-Object -Property DriveType -EQ -Value Removable;

        #Check if we found multiple removable drives
        if ($Script:Volume.Count -gt 1)
        {
            Write-Debug "Multiple removeable drives found";
            Write-Host "`nMultiple removeable drives found" -ForegroundColor Red -BackgroundColor Black;
            Write-Host "`nAs a safety precaution when multiple USB Drives are found you must explicitly pass the drive letter to this script.";
            Write-Host "e.g. .\$($Script:ScriptName).ps1 -DriveLetter A`n`n";
            Break;
        }
    }
    else
    {
        Write-Verbose "Getting the volume object for the drive $DriveLetter";
        $Script:Volume = Get-Volume -DriveLetter "$Script:DriveLetter" -ErrorAction SilentlyContinue | Where-Object -Property DriveType -EQ -Value Removable;
    }

    if ($Script:Volume.DriveLetter -ne $null)
    {
        Write-Host "Removable USBDrive found" -ForegroundColor Green;
        @{"DriveLetter"="$($Script:Volume.DriveLetter):";"Label"="$($Script:Volume.FileSystemLabel)"} | Format-Table -HideTableHeaders;
        
        #check we have the right volume
        $title = "Format and Configure USB drive"
        $message = "Do you want to format and configure volume $($Script:Volume.DriveLetter): ($($Script:Volume.FileSystemLabel))?"
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Formats the volume and copies the contents from $($Script:CustomerTools)."
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Exits and does not perform any action on the volume"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

        #check what values the user entered
        switch ($result)
        {
            0 #Yes
            {
                Write-Verbose "Choice = Yes";
                
                # Display the progress bar
                $PgressBarPerComp = 0;
                $PgressBarStep = 1;
                Write-Progress -Activity "Preparing USB Drive" -status "Step $PgressBarStep of 3 -> Formating Drive -> Elapsed Time: $(FormatElapsedTime $Script:ElapsedTime.Elapsed)" -PercentComplete $PgressBarPerComp;

                #Format the volume
                Write-Verbose "Formating volume";
                #Format-Volume -ObjectId $Volume.ObjectId -DriveLetter $Volume.DriveLetter -NewFileSystemLabel $Volume.FileSystemLabel -FileSystem $Volume.FileSystem -Force;
                $Result = Format-Volume -ObjectId $($Volume.ObjectId) -FileSystem $($Script:Volume.FileSystem) -NewFileSystemLabel $($Script:Volume.FileSystemLabel) -Force

                #complete the progress bar to prevent it displaying during user input
                Write-Progress -Completed $true;

                <#
                bitlocker - Start
                #>

                #Get a password for bitlocker encryption (WarningAction=SilentlyContinue to supress warning regarding backup key from user)
                Write-Verbose "Prompting for bitlocker password";
                Write-Host "`n`nBitLocker2Go Setup" -ForegroundColor Green;
                $BitLockerPassword = Read-Host -Prompt "Enter a password to encrypt your USB with BitLocker2Go" -AsSecureString;

                # Display the progress bar
                $PgressBarPerComp = 0;
                $PgressBarStep = 2;
                Write-Progress -Activity "Preparing USB Drive" -status "Step $PgressBarStep of 3 -> BitLocker Encryption -> Elapsed Time: $(FormatElapsedTime $Script:ElapsedTime.Elapsed)" -PercentComplete $PgressBarPerComp;

                #generate the recovery key
                Write-Verbose "Generating a Recovery Password Key for the drive (GPO requirement and Best Practice)";
                $Result = Add-BitLockerKeyProtector -MountPoint "$($Script:Volume.DriveLetter):" -RecoveryPasswordProtector -WarningAction SilentlyContinue;

                #Save the recovery password key to a file
                Write-Verbose "Saving the Recovery Password Key to a file share";
                "Bitlocker Key for $($Script:Volume.FileSystemLabel)`r`n `
                Identifier: $((Get-BitLockerVolume "$($Script:Volume.DriveLetter):").KeyProtector.KeyProtectorId)`r`n `
                Key: $((Get-BitLockerVolume "$($Script:Volume.DriveLetter):").KeyProtector.RecoveryPassword)" | Out-File -FilePath "BitLockerKeys:\$($Script:Volume.FileSystemLabel).BitLockerKey.txt";
                Write-Host "Important: " -ForegroundColor Green -NoNewline;
                Write-Host "BitLocker Key saved to export location";

                #Enable bitlocker on the drive (WarningAction=SilentlyContinue to supress warning regarding backup key from user)
                Write-Verbose "Enabling Bitlocker on the drive";
                $Result = Enable-BitLocker -MountPoint "$($Script:Volume.DriveLetter):" -EncryptionMethod Aes256 -UsedSpaceOnly -Password $BitLockerPassword -PasswordProtector -WarningAction SilentlyContinue;

                #Wait for encryption to complete
                Do 
                {
                    $PgressBarPerComp = (Get-BitLockerVolume -MountPoint "$($Script:Volume.DriveLetter):").EncryptionPercentage;
                    Write-Verbose "Waiting for BitLocker Encryption, current progress $($PgressBarPerComp)%";
                    Write-Progress -Activity "Preparing USB Drive" -status "Step $PgressBarStep of 3 -> BitLocker Encryption -> Elapsed Time: $(FormatElapsedTime $Script:ElapsedTime.Elapsed)" -PercentComplete $PgressBarPerComp;
                    Start-Sleep -Milliseconds 100;
                }
                until  ($PgressBarPerComp -eq 100)

                <#
                bitlocker - End
                #>

                if (Test-Path -Path $Script:CustomerTools)
                {
                    #Get the collection of files to copy
                    $FilesToCopy = Get-ChildItem -Path $CustomerTools -Recurse;

                    #Count the number of items to copy
                    $CustomerToolsCnt = $FilesToCopy.Count
                    $PgressBarStep = 3;
                    $CustomerToolsIdx = 0;

                    #copy the customer tools
                    Write-Verbose "Copying tools directory $($Script:CustomerTools)";
                    #$CopyJob = Start-Job -ScriptBlock {Copy-Item -Path $CustomerTools -Destination "$($Volume.DriveLetter):\" -Recurse};
                    #do
                    #{
                    #    $CopyJob.Progress | fl *
                    #}
                    #until ($CopyJob.State -eq "Completed")
                    #Copy-Item -Path $CustomerTools -Destination "$($Volume.DriveLetter):\" -Recurse -passthru | Write-Progress -Activity "Preparing USB Drive" -status "Step $PgressBarStep of 3 -> Copying files -> Elapsed Time: $(FormatElapsedTime $ElapsedTime.Elapsed)" -PercentComplete ((($CustomerToolsIdx++) / $CustomerToolsCnt)*100);                 
                    #$BITSTransfer = Start-BitsTransfer -Source $CustomerTools -Destination "$($Volume.DriveLetter):\" -Asynchronous;
                    Foreach ($file in $FilesToCopy)
                    {
                        Write-Progress -Activity "Preparing USB Drive" -status "Step $PgressBarStep of 3 -> Copying files -> Elapsed Time: $(FormatElapsedTime $Script:ElapsedTime.Elapsed)" -PercentComplete ((($CustomerToolsIdx++) / $CustomerToolsCnt)*100);                 
                        $sourceFile = $file.FullName;
                        $destFile = Join-Path -Path "$($Script:Volume.DriveLetter):\" -ChildPath $file.FullName.Replace($Script:CustomerTools.Replace("*",""),"");
                        Write-Verbose "Source File: $sourceFile";
                        Write-Verbose "Dest File: $destFile";
                        #make sure the parent destination directory exists
                        if ((Test-Path (Split-Path $destFile -Parent)) -eq $false) {$Result = New-Item -Path (Split-Path $destFile -Parent) -ItemType Directory};
                        #copy the file
                        Copy-Item -Path $sourceFile -Destination $destFile;
                    }

                }

                Write-Verbose "Done";
                Write-Host "Complete: " -ForegroundColor Green -NoNewline;
                Write-Host "Volume $($Script:Volume.DriveLetter) ($($Script:Volume.FileSystemLabel)) has been formated, encrypted and prepared for Customer visit";
            }
    
            1 #No
            {
                Write-Verbose "Choice = No";
                Write-Host "You selected 'No', script exiting and no action will be taken" -ForegroundColor Yellow -BackgroundColor Black;
            }
        }
    }
    else
    {
        Write-Verbose "Volume Count = $($Script:Volume.Count)";
        Write-Host "There were $($Script:Volume.Count) removable volumes found. Unable to proceed." -ForegroundColor Red -BackgroundColor Black;
    }
}

end
{
    #remove the PSDrive
    Write-Verbose "Removing the PSDrive";
    Remove-PSDrive -Name BitLockerKeys;

    #complete the progress bar if not already closed
    Write-Progress -Completed $true;

    # Stop the timer
    $Script:ElapsedTime.Stop()
    $Script:ElapsedTime.Reset()

    # pop back to the original location
    Pop-Location;
}