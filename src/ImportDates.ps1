<#
This script imports dates from a delimited file into OpCon.
You can use traditional MSGIN functionality or the OpCon API.  Also added a "debug" option
so that you can view the dates that will be added.
Author: Bruce Jernell
Version: 1.0
#>
param(
    [string]$opconmodule,                                             # Path to OpCon API functions module
    [string]$msginPath = "C:\ProgramData\OpConxps\MSLSAM\MSGIN",      # Path to MS LSAM MSGIN directory
    [string]$url,                                                     # OpCon API URL ie https://<opcon>
    [string]$apiUser,                                                 # OpCon API User
    [string]$apiPassword,                                             # OpCon API Password
    [string]$extUser,                                                 # OpCon External Event user
    [string]$extPassword,                                             # OpCon External Event password
    [string]$extToken,                                                # OpCon External Token (OpCon Release 20+)
    [Parameter(Mandatory=$true)] [string]$filePath,                                                # Path to calendar dates path
    [Parameter(Mandatory=$true)] [string]$fileDelimiter,                                           # Delimiter in dates file ie ","
    [Parameter(Mandatory=$true)] [string]$calendar,                                                # OpCon Calendar (ex "Master Holiday")
    [string]$option = "debug"                                         # Script option: "api", "msgin", "debug"
)

$ErrorActionPreference = "stop"

if(Test-Path $filePath)
{
    $datesArray = Get-Content $filePath -Raw
    $datesArray.Split("`n") | ForEach-Object{ 
        if($_.Trim())
        {   $dateString = $dateString + $_.Trim() + $fileDelimiter } 
    }
    $dateString = $dateString.TrimEnd($fileDelimiter).Replace($fileDelimiter,";")
    if($dateString)
    {
        Switch ($option)
        {
            "api"   {
                #Verifies opcon module exists and is imported
                if(Test-Path $opconmodule)
                {
                    #Verify PS version is at least 3.0
                    if($PSVersionTable.PSVersion.Major -ge 3)
                    {
                        #Import needed module
                        Import-Module -Name $opconmodule -Force      
                    }
                    else
                    {
                        Write-Output "Powershell version needs to be 3.0 or higher!"
                        Exit 100
                    }
                }
                else
                {
                    Write-Output "Unable to import OpCon API module!"
                    Exit 100
                }

                #Force TLS 1.2
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

                #Skip self signed certificates (OpCon API default)
                OpCon_SkipCerts

                if($extToken)
                { $token = "Token " + $extToken }
                else
                { $token = "Token " + (OpCon_Login -url $url -user $apiUser -password $apiPassword).id }

                OpCon_UpdateCalendar -url $url -token $token -name $calendar -date $dateString

                if($error)
                { 
                    Write-Output $error
                    Write-Output "There was a problem updating calendar"$calendar" with date "$dateString 
                    Exit 5
                }
                break
            }
            "msgin" {
                if($msginPath)
                {
                    if(test-path $msginPath)
                    {   Write-Output "$msginPath path exists" }
                    else
                    {
                        Write-Output "$msginPath path does not exist"
                        Exit 101
                    }
                }
                else
                {
                    Write-Output "MSGIN Path parameter must be specified!"
                    Exit 102
                }    

                if($extToken)
                { $extPassword = $extToken }

                # Creates MSGIN file
                Write-Output ("Sending date/s to calendar $calendar via MSGIN:`n" + $dateString)
                "`$CALENDAR:ADD,$calendar," + $dateString + ",$extUser,$extPassword" | Out-File -FilePath ($msginPath + "\events.txt") -Encoding ascii
                break
            }
            "debug" {
                Write-Output ("Date/s to add:`n" + $dateString); break
            }
            default {
                Write-Output "Invalid option, must use 'api','msgin', or 'debug'."
                Exit 1
            }
        }
    }
    else{
        Write-Output "Issue parsing dates from $filePath with delimiter $fileDelimiter"
        Exit 3
    }
}
else{
    Write-Output "Invalid filePath specified."
    Exit 2
}