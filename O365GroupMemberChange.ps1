#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####
# Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
# Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
if (!(Test-Path $scriptCSV))
{
    ######### Template
    "GroupEmail,MemberEmail,AddRemove" | Add-Content $scriptCSV
    "mygroup@contoso.com,user1@contoso.com,Add" | Add-Content $scriptCSV
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
PressEnterToContinue
$no_errors = $true
$error_txt = ""
$results = @()
# region Connect to M365
$myscopes=@()
$myscopes+="User.ReadWrite.All"
$myscopes+="GroupMember.ReadWrite.All"
$myscopes+="Group.ReadWrite.All"
$connected_ok = ConnectMgGraph $myscopes
# endregion
if (-not ($connected_ok))
{
    Write-Host "Connection failed."
}
else
{ # M365 Connected
    Write-Host "--------------------"
    $processed=0
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
            #######
            ####### Start code for object $x
            #region Object X
            #X:GroupEmail,MemberEmail,AddRemove
            <#                 
            Mg-graph Cannot Update a mail-enabled security groups and or distribution list.
            https://learn.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0&tabs=http
            Get-DistributionGroup
            Get-UnifiedGroup
            Get-AzureADGroup Security group
            #>
            $user = Get-MgUser -UserId $x.MemberEmail
            if (-not $user)
            { # user bad
                Write-Host "User not found: $($x.Mail) ERR"  -ForegroundColor Red
            } # user bad
            else
            { # user ok
                $group = Get-MgGroup -Filter "mail eq '$($x.GroupEmail)'"
                if (-not $group) 
                { # group bad
                    Write-Host "Group not found: $($x.GroupEmail) ERR"  -ForegroundColor Red
                } # group bad
                else
                { # group ok
                    $isMember = Get-MgGroupMember -GroupId $group.Id | Where-Object { $_.Id -eq $user.Id }
                    If ($x.AddRemove -eq "Add")
                    { # Add
                        if ($isMember) {
                            Write-Host "User already in group. OK" -ForegroundColor Yellow
                        } else {
                            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.id
                            Write-Host "User added to group. OK" -ForegroundColor Green
                        }
                    } # Add
                    Elseif ($x.AddRemove -eq "Remove")
                    { # Remove
                        if (-not $isMember) {
                            Write-Host "User already removed from group. OK" -ForegroundColor Yellow
                        } else {
                            Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id 
                            Write-Host "User removed from group. OK" -ForegroundColor Green
                        }
                    } # Remove
                    Else
                    {
                        Write-Host "AddRemove column has invalid data (should be Add or Remove): $($x.AddRemove) ERR"  -ForegroundColor Red
                    }
                } # group ok
            } # user ok
            #endregion Object X
            ####### End code for object $x
            #######
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
PressEnterToContinue



<# 
####### Start code for object $x
$GroupEmail = $x.GroupEmail.Trim()
$MemberEmail =$x.MemberEmail.Trim()
$AddRemove = $x.AddRemove.trim()
##########
$no_err=$true
$group = $null
##########
$ident = $GroupEmail
$group_type = ""
if (-not $group)
{
    Try {$group = Get-DistributionGroup -identity $ident -ErrorAction SilentlyContinue}
    Catch {}
    if ($group) {$group_type = "DistributionGroup"}
}
if (-not $group)
{
    $ident = $GroupEmail
    Try {$group = Get-UnifiedGroup -identity $ident -ErrorAction SilentlyContinue}
    Catch {}
    if ($group) {$group_type = "UnifiedGroup"}
}
if (-not $group)
{

    ##
    Try {$group = Get-AzureADGroup | Where-Object -Property "DisplayName" -EQ $ident -ErrorAction SilentlyContinue}
    Catch {}
    if ($group) {$group_type = "SecurityGroup"}
}
###########
if ($group_type -eq "")
{
    $no_err=$false
    Write-Warning "$($ident) is not a known group"
}
if ($group_type -eq "UnifiedGroup")
{
    $links= @($group|Get-UnifiedGroupLinks -LinkType Member |Select-Object WindowsLiveID|Sort-Object WindowsLiveID)
    $members= $links.WindowsLiveID
    $memberemails= $members -join ", "
    $membersabbrev= $members.Replace($group_domain,"")
    ###
    Write-Host "$($group_type) Prior Members: $($membersabbrev)"
    if ($AddRemove -eq "Remove")
        {### Remove
		    ##########
	        $ident = $MemberEmail
	        Try {$recip = Get-Recipient -identity $ident;$MemberEmail=$recip.PrimarySmtpAddress}
	        Catch {Write-Warning "$($ident) is not a known email, trying to remove anyway"}
	        ##########
            if ($Members.Contains($MemberEmail))
            {
                #Remove-DistributionGroupMember $group.PrimarySmtpAddress -member $MemberEmail -BypassSecurityGroupManagerCheck -Confirm:$false
                Remove-UnifiedGroupLinks  -Identity $group.PrimarySmtpAddress -LinkType Members -Links $MemberEmail -Confirm:$false
                Write-Host "OK: Removed"
            }
            else
            {
                Write-Host "OK: Wasn't an Member"
		    }
        }### Remove
        Else
        { ### Add
	        ##########
	        $ident = $MemberEmail
	        Try {$recip = Get-Recipient -identity $ident -ErrorAction SilentlyContinue}
	        Catch {Write-Warning "$($ident) is not a known email";$no_err=$false}
            if (-not $recip)
            {
                Write-Warning "$($ident) is not a known email";$no_err=$false
            }
	        ##########
	        if ($no_err)
		    {
		        if ($Members.Contains($MemberEmail))
                {
                
                    Write-Host "OK: Already a member"
                }
                else
                {
                    #Add-DistributionGroupMember $group.PrimarySmtpAddress -member $MemberEmail -BypassSecurityGroupManagerCheck
                    Add-UnifiedGroupLinks -Identity $group.PrimarySmtpAddress -LinkType Members -Links $MemberEmail
                    Write-Host "OK: Added"
		        }
		    }
        } ### Add
    ###
}
if ($group_type -eq "DistributionGroup")
{
    $group_domain = ($group.PrimarySmtpAddress -split "@")[1]
    $members= @(Get-DistributionGroupMember -Identity $group.PrimarySmtpAddress -ResultSize Unlimited | Select-Object -ExpandProperty PrimarySmtpAddress | Sort-Object)
    $memberemails= $members -join ", "
    $membersabbrev= $memberemails.Replace($group_domain,"")
    ###
    Write-Host "$($group_type) Prior Members: $($membersabbrev)"
    if ($AddRemove -eq "Remove")
    {### Remove
		##########
	    $ident = $MemberEmail
	    Try {$recip = Get-Recipient -identity $ident;$MemberEmail=$recip.PrimarySmtpAddress}
	    Catch {Write-Warning "$($ident) is not a known email, trying to remove anyway"}
	    ##########
        if ($Members.Contains($MemberEmail))
        {
            Remove-DistributionGroupMember $group.PrimarySmtpAddress -member $MemberEmail -BypassSecurityGroupManagerCheck -Confirm:$false
            Write-Host "OK: Removed"
        }
        else
        {
            Write-Host "OK: Wasn't an Member"
		}
    }### Remove
    Else
    { ### Add
	    ##########
	    $ident = $MemberEmail
	    Try {$recip = Get-Recipient -identity $ident -ErrorAction SilentlyContinue}
	    Catch {Write-Warning "$($ident) is not a known email";$no_err=$false}
        if (-not $recip)
        {
            Write-Warning "$($ident) is not a known email";$no_err=$false
        }
	    ##########
	    if ($no_err)
		{
		    if ($Members.Contains($MemberEmail))
            {
                
                Write-Host "OK: Already a member"
            }
            else
            {
                Add-DistributionGroupMember $group.PrimarySmtpAddress -member $MemberEmail -BypassSecurityGroupManagerCheck
                Write-Host "OK: Added"
		    }
		}
    } ### Add
}
if ($group_type -eq "SecurityGroup")
{
    #$group_domain = ($group.PrimarySmtpAddress -split "@")[1]
    $members= @($group | Get-AzureADGroupMember)
    $memberobjectids= @($members.ObjectId | Sort-Object)
    $memberemails= ($members.UserPrincipalName | Sort-Object) -join ", "
    $membersabbrev= $memberemails
    #$membersabbrev= $memberemails.Replace($group_domain,"")
    ###
    Write-Host "$($group_type) Prior Members: $($membersabbrev)"
    if ($AddRemove -eq "Remove")
    {### Remove
		##########
	    $ident = $MemberEmail
	    Try {$recip = Get-AzureADUser -ObjectId $ident;$MemberEmail=$recip.ObjectId}
	    Catch {Write-Warning "$($ident) is not a known email, trying to remove anyway"}
	    ##########
        if ($memberobjectids.Contains($recip.ObjectId))
        {
            #Remove-DistributionGroupMember $group.PrimarySmtpAddress -member $MemberEmail -BypassSecurityGroupManagerCheck -Confirm:$false
            Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $recip.ObjectId
            Write-Host "OK: Removed"
        }
        else
        {
            Write-Host "OK: Wasn't an Member"
		}
    }### Remove
    Else
    { ### Add
	    ##########
	    $ident = $MemberEmail
	    Try {$recip = Get-AzureADUser -ObjectId $ident -ErrorAction SilentlyContinue}
	    Catch {Write-Warning "$($ident) is not a known email";$no_err=$false}
        if (-not $recip)
        {
            Write-Warning "$($ident) is not a known email";$no_err=$false
        }
	    ##########
	    if ($no_err)
		{
		    if ($memberobjectids.Contains($recip.ObjectId))
            {
                
                Write-Host "OK: Already a member"
            }
            else
            {
                #Add-DistributionGroupMember $group.PrimarySmtpAddress -member $MemberEmail -BypassSecurityGroupManagerCheck
                Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $recip.ObjectId
                Write-Host "OK: Added"
		    }
		}
    } ### Add
}
####### End code for object $x #>