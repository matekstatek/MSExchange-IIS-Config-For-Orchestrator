<#
.Synopsis
   Sets attributes in IIS.
.DESCRIPTION
   The script sets attribute PSLanguageMode to FullLanguage in Exchange Back End/Powershell and 
   Default Web Site/Powershell. An user should run it without any attachements, but in case of error 
   the script creates csv file with backup of previous configuration.
.EXAMPLE
   .\Set-IISForOrchestrator.ps1
.EXAMPLE
   .\Set-IISForOrchestrator.ps1 -inputFileName .\EBEConfig_backup.csv
.EXAMPLE
   .\Set-IISForOrchestrator.ps1 -inputFileName .\EBEConfig_backup.csv -WhatIf
#>

[cmdletbinding(SupportsShouldProcess=$True)]
param
(
    [Parameter(Mandatory = $false)]
    [string]$inputFileName
)

function Write-DWSConfig
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param()

    if($dwsIsExisting)
    {
        write-host (Get-Date -Format "hh:mm:ss").ToString() "[WARNING] The specified attribute in 'Default Web Site\PowerShell' is already exist." -ForegroundColor Yellow
    }

    else
    {
        Add-WebConfigurationProperty -pspath $dws_path -filter "/appSettings" -AtIndex 0 -Name "Collection" -Value @{key="PSLanguageMode";value="FullLanguage"}
        write-host (Get-Date -Format "hh:mm:ss").ToString() "[INFO] The specified attribute in 'Default Web Site\PowerShell' was successfully added." -ForegroundColor Green
    }
}

# save config to file as backup
function Save-ToFile
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param()

    $currentConfig = @()

    for($i=0; $i -lt $up.Length; $i++)
    {
        $item = New-Object PSObject
        $item | 
            Add-Member -type NoteProperty -Name 'Key'   -Value (  $up[$i]  )
        $item | 
            Add-Member -type NoteProperty -Name 'Value' -Value ( $down[$i] )

        $currentConfig += $item
    }

    $currentConfig | 
        Export-Csv -Path $SavedFileName -NoTypeInformation -Force
}

# change backend config
function Write-EBEConfig
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param()

    $attributeAlreadyExists = $false

    if($inputFileName)
    {
        Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[INFO] The file has been attached. The configuration will be taken from the file $inputFileName." -ForegroundColor DarkGreen

        # if file is not exist
        if(![System.IO.File]::Exists($inputFileName))
        {
            Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[ERROR] The file does not exist. Pick another file." -ForegroundColor Red

            throw [System.IO.FileNotFoundException]
            return;
        }

        # if attached file is empty
        if((Get-Content $inputFileName) -eq $Null)
        {
            Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[ERROR] The file is empty. Pick another file." -ForegroundColor Red

            throw [System.IO.FileLoadException]
            return;
        }

        # to do:
        # if file is not in right format

        # if everything is ok
        else
        {
            $file = Get-Content $inputFileName

            $up = @()
            $down = @()

            foreach ($line in $file)
            {
                if ($line[0] -eq '"' -and $line -notcontains '"Key","Value"')
                {
                    $splittedLine = $line.Split(',') -replace '"'
                    $up   += $splittedLine[0]
                    $down += $splittedLine[1]
                } 
            }

            write-host (Get-Date -Format "hh:mm:ss").ToString() "[INFO] The file loaded successfully." -ForegroundColor DarkGreen
        }
    }
    
    # getting data from exchange backend
    else
    {
        write-host (Get-Date -Format "hh:mm:ss").ToString() "[INFO] Configuration based on IIS data. In case of error, run the script by attaching the config file $SavedFileName located in the script location." -ForegroundColor DarkGreen
        
        $up = (get-WebConfigurationProperty -pspath $ebe_path -filter "/appSettings" -name "collection" | 
            select key).Key
        $down = (get-WebConfigurationProperty -pspath $ebe_path -filter "/appSettings" -name "collection" | 
            select value).Value
        
        for($i=0; $i -lt $up.Count; $i++)
        {
            if ($up[$i] -eq "PSLanguageMode")
            {
                if ($down[$i] -eq "FullLanguage")
                {
                    write-host (Get-Date -Format "hh:mm:ss").ToString() "[WARNING] Attribute Exchange Back End\Powershell is correct. Nothing to do here." -ForegroundColor Yellow

                    $attributeAlreadyExists = $true
                }

                else
                {
                    $down[$i] = "FullLanguage"
                }
            }
        }
    }

    if($up.Length -ne 0 -and -not $attributeAlreadyExists)
    {
        for($i=0; $i -lt $up.Count; $i++)
        {
            if ($up[$i] -eq "PSLanguageMode")
            {
                $down[$i] = "FullLanguage"
            }
        }

        Remove-WebConfigurationProperty -pspath $ebe_path -filter "/appSettings" -name "Collection"

        for($i=0; $i -lt $up.Length; $i++)
        {
            Write-Progress -Activity "Creating attributes in Exchange Back End" -Status "$([Math]::Round($(100*$i/$up.Length),2))% Complete" -PercentComplete $(100*$i/$up.Length)
            Add-WebConfigurationProperty -pspath $ebe_path -filter "/appSettings" -Value @{key=$up[$i];value=$down[$i]} -name "Collection"
        }

        Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[INFO] The specified attribute in 'Exchange Back End\PowerShell' was successfully added." -ForegroundColor Green
    }
}

# end function
function Get-Result
{
    $EBE = Get-WebConfigurationProperty -pspath $ebe_path -filter "/appSettings" -name "collection" 

    if ($EBE)
    {
        Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[INFO] Script successfully completed."                                              -ForegroundColor Green
    }

    else
    {
        Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[ERROR] An error occured while writing attribute to Exchange Back End\Powershell."  -ForegroundColor Red
        Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[ERROR] Run script with config file attached."                                      -ForegroundColor Red
    }

    Write-Host (Get-Date -Format "hh:mm:ss").ToString() "[INFO] The attributes are as follows:`n"
    Write-Host "Default Web Site:" -BackgroundColor DarkCyan

    Get-WebConfiguration -pspath $dws_path -filter "/appSettings/add" | 
        select key, value | 
            ft -a | 
                Out-String | 
                    Write-Host

    Write-Host "`n`nExchange Back End:" -BackgroundColor DarkCyan

    $EBE | 
        select key, value | 
            ft -a | 
                Out-String | 
                    Write-Host
}

# dws : Default Web Site
# ebe : Exchange Back End
$dws_path = "IIS:\Sites\Default Web Site\PowerShell\"
$ebe_path = "IIS:\Sites\Exchange Back End\PowerShell\"


if ($inputFileName -ne "")
{
    $inputFileName = (Get-Location).ToString() + $inputFileName
}

$SavedFileName = (Get-Location).ToString() + "\EBEConfig_backup_" + (Get-Date -Format 'dd.MM.yy_hh.mm.ss').ToString() + ".csv"
$dwsIsExisting = (Get-WebConfiguration -pspath $dws_path -filter "/appSettings/add[@key='PSLanguageMode']")

try
{
    Save-ToFile
    Write-DWSConfig
    Write-EBEConfig

    iisreset.exe

    Get-Result
}
catch 
{
    return
}