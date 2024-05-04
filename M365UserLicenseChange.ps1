#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####
# Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
# 
# Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
#
if (!(Test-Path $scriptCSV))
{
    ######### Template
    "User,LicensesToAdd,LicensesToRemove" | Add-Content $scriptCSV
    "myuser@contoso.com,SPB," | Add-Content $scriptCSV
    ######### 
	$ErrOut=201; Write-Host "Err $ErrOut : Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. Edit CSV and run again.";Pause; Exit($ErrOut)
}
# ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV)
$entriescount = $entries.count
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in O365"
Write-Host ""
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV -leaf) ($($entriescount) entries)"
$entries | Format-Table
Write-Host "-----------------------------------------------------------------------------"
Read-host "Press <Enter> to continue"
$no_errors = $true
$error_txt = ""
$results = @()
# region Connect to M365
$myscopes=@()
$myscopes+="User.ReadWrite.All"
$myscopes+="GroupMember.ReadWrite.All"
$myscopes+="Group.ReadWrite.All"
$connectedall_ok = ConnectMgGraph $myscopes
# endregion
if ($connectedall_ok)
{ # M365 Connected
    Write-Host "--------------------"
    $processed=0
    $message="$entriescount Entries. Continue?"
    $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","&No")
    [int]$defaultChoice = 0
    $choiceRTN = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
    if ($choiceRTN -eq 1)
    { "Aborting" }
    else 
    { # continue choices
        $choiceLoop=0
        $i=0        
        foreach ($x in $entries)
        { # each entry
            $i++
            write-host "-----" $i of $entriescount $x
            if ($choiceLoop -ne 1)
            { # Process all not selected yet, Ask
                $message="Process entry "+$i+"?"
                $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","Yes to &All","&No","No and E&xit") # 0 Yes, 1 All, 2 No, 3 Exit
                [int]$defaultChoice = 1
                $choiceLoop = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
            } # Process all not selected yet, Ask
            if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
            { # Process
                $processed++
                ####### Start code for object $x
                $user = Get-MgUser -UserId $x.User
                if ($user)
                { # user ok
                    $SubscribedSkus = Get-MgSubscribedSku -All
                    #$SubscribedSkus| Select-Object SkuPartNumber, ConsumedUnits, @{N="Prepaid";E={$_.PrepaidUnits.Enabled}}  | Format-Table | Out-String | Write-Host
                    $SkuPartNumbersToAdd    = @()
                    $SkuPartNumbersToAdd    += $x.LicensesToAdd.Split(",").trim() | Where-Object {$_ -ne ""}
                    $SkuPartNumbersToRemove = @()
                    $SkuPartNumbersToRemove += $x.LicensesToRemove.Split(",").trim() | Where-Object {$_ -ne ""} | Where-Object {$_ -ne "<all>"}
                    $SkuPartNumbersToTest = @()
                    $SkuPartNumbersToTest += $SkuPartNumbersToAdd
                    $SkuPartNumbersToTest += $SkuPartNumbersToRemove
                    $skusok = $true
                    ForEach ($SkuPartNumberToTest in $SkuPartNumbersToTest)
                    { # test each sku
                        if ($SkuPartNumberToTest -notin $SubscribedSkus.SkuPartNumber)
                        { # sku bad
                            Write-Host "  User: $($user.DisplayName) - " -NoNewline
                            Write-Host "Sku not found: $($SkuPartNumberToTest) ERR" -ForegroundColor Red
                            Read-Host "Press <Enter> to see a list of valid SKUs"
                            $SubscribedSkus | Sort-Object SkuPartNumber `
                            | Select-Object SkuPartNumber, ConsumedUnits, @{N="Prepaid";E={$_.PrepaidUnits.Enabled}} `
                            | Select-Object SkuPartNumber,Prepaid,ConsumedUnits, @{N="Availabled";E={$_.Prepaid - $_.ConsumedUnits}} `
                            | Format-Table | Out-String | Write-Host
                            Read-Host "Press <Enter> to continue"
                            $skusok = $false
                            Break # break out of for loop
                        } # sku bad
                    } # test each sku
                    if ($skusok)
                    { # sku ok
                        $userskus=@((Get-MgUserLicenseDetail -UserId $user.id).SkuPartNumber| Where-Object {$_ -ne ""}| Where-Object {$_ -ne $null})
                        if ($x.LicensesToRemove -eq "<all>")
                        { # wants to remove all licenses
                            $SkuPartNumbersToRemove = @($userskus)
                        } # wants to remove all licenses
                        # If something is in add AND remove, add should win
                        $SkuPartNumbersToRemove = @($SkuPartNumbersToRemove | Where-Object {$_ -NotIn $SkuPartNumbersToAdd})
                        Write-Host "  User: $($user.DisplayName) [$($userskus -join ", ")] - " -NoNewline
                        ForEach ($SkuPartNumberToAdd in $SkuPartNumbersToAdd)
                        { # test each sku
                            if ($SkuPartNumberToAdd -notin $userskus)
                            { # sku bad
                                $skusok = $false
                                Break # break out of for loop
                            } # sku bad
                        } # test each sku
                        if ($skusok)
                        { # skusok for add side but test remove side
                            ForEach ($SkuPartNumberToRemove in $SkuPartNumbersToRemove)
                            { # test each sku
                                if ($SkuPartNumberToRemove -in $userskus)
                                { # sku bad
                                    $skusok = $false
                                    Break # break out of for loop
                                } # sku bad
                            } # test each sku
                        } # skusok for add side but test remove side
                        if ($skusok)
                        { # user has skus
                            Write-Host "User licenses already OK" -ForegroundColor Yellow
                        } # user has skus
                        else 
                        { # user needs sku change
                            # get an array of SkuIds
                            $SkusToAdd = $SubscribedSkus | Where-Object SkuPartNumber -in $SkuPartNumbersToAdd | Select-Object SkuId -ExpandProperty SkuId
                            $SkusToRemove = $SubscribedSkus | Where-Object SkuPartNumber -in $SkuPartNumbersToRemove | Select-Object SkuId -ExpandProperty SkuId
                            # AddLicenses needs an array of hashvalues {SkuId='xxxx-xxxx'}
                            $SkusToAddHashArray = @()
                            ForEach ($SkuToAdd in $SkusToAdd)
                            {
                                $SkusToAddHashArray += @{SkuId = $SkuToAdd.SkuId}
                            }
                            $ret= Set-MgUserLicense -UserID $user.id -AddLicenses $SkusToAddHashArray -RemoveLicenses @($SkusToRemove)
                            if($?)
                            { # command succeeded
                                $userskus_after =(Get-MgUserLicenseDetail -UserId $user.id).SkuPartNumber
                                Write-Host " changed to [$($userskus_after -join ", ")] " -NoNewline
                                Write-Host "Licenses changed OK" -ForegroundColor Green
                            }
                            else
                            {# command failed
                                Write-Host "Something went wrong ERR" -ForegroundColor Yellow
                                Write-Host $Error[0].Exception.Message
                                Read-Host "Press <Enter>"
                            }
                        } # user needs sku change
                    } # sku ok
                } # user ok
                else
                { # no user
                    Write-Host "User not found: $($x.Mail) ERR"  -ForegroundColor Red
                } # no user
                ####### End code for object $x
            } # Process
            if ($choiceLoop -eq 2)
            {
                write-host ("Entry "+$i+" skipped.")
            }
            if ($choiceLoop -eq 3)
            {
                write-host "Aborting."
                break
            }
        } # each entry
    } # continue choices
    WriteText "Removing any open sessions..."
    Get-PSSession 
    Get-PSSession | Remove-PSSession
    WriteText "------------------------------------------------------------------------------------"
    $message ="Done. " +$processed+" of "+$entriescount+" entries processed. Press [Enter] to exit."
    WriteText $message
    WriteText "------------------------------------------------------------------------------------"
	# Transcript Save
    Stop-Transcript | Out-Null
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
    $TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
    If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
    Move-Item $Transcript $TranscriptTarget -Force
    # Transcript Save
} # M365 Connected
PauseTimed -quiet 3 #$message
Pause