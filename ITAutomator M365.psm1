########################################################################
# ITAutomator M365.psm1 copyright(c) ITAutomator
# https://www.itautomator.com
# 
# Library of useful functions for PowerShell Programmers.
# These functions are useful for Microsoft 365 programming.
#
#################################################################

<#
Usage: Include this at the top of your main procedure
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
#>

<#
Version History
2024-05-06
ConnectExchangeOnline
ConnectMgGraph
GroupChildren - Returns the children (users) of this group or role
GroupParents - Returns the parents of this user (groups that this user is in)
MgGroupCreate - Creates a group in Azure/Entra (if necessary)

2024-05-03 Initial Version
MgGroupCreate

Alphabetical list of functions
-----------------------------------
ConnectExchangeOnline
ConnectMgGraph
GroupChildren - Returns the children (users) of this group or role
GroupParents - Returns the parents of this user (groups that this user is in)
MgGroupCreate - Creates a group in Azure/Entra (if necessary)
-----------------------------------
#>
Function ConnectExchangeOnline ($domain=$null)
{
    do # connect loop
    {
        $connectedall_ok = $true
        ###############
        Write-Host "Connect-ExchangeOnline" -ForegroundColor Yellow
        # PressEnterToContinue
        # Load the module and show results
        $module= "ExchangeOnlineManagement" ; Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
        $connected_ok=$false
        # Check if there are any Exchange Online sessions
        $exchangeSession = Get-ConnectionInformation
        if ($exchangeSession) {
            Write-Output "Already connected to Exchange Online. (Use Disconnect-ExchangeOnline to change domains)"
            $connected_ok=$true
        }
        if (-not $connected_ok)
        {Connect-ExchangeOnline -ShowBanner:$false}
        $domain_exch = Get-AcceptedDomain | Where-Object Default -eq $true| Select-object -ExpandProperty DomainName
        Write-Host "Domain :" -NoNewline
        Write-Host $domain_exch -ForegroundColor Green
        if ($domain)
        { # we need it to be this domain
            if ($domain -ne $domain_exch)
            {
                Write-Host "Domains don't match. Using Disconnect-ExchangeOnline and trying again."
                PressEnterToContinue
                Disconnect-ExchangeOnline -Confirm:$false
                Continue #start loop again
            }
        } # we need it to be this domain
        ###############
        if ($connectedall_ok)
        {Break}
    } while ($true) # connect loop
    Return $connectedall_ok
}
Function ConnectMgGraph ($myscopes=$null, $domain=$null)
{
    $junk = $null # collect throwaway output
    # Load the module and show results
    $module= "Microsoft.Graph.Authentication" ; Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
    if (-not $myscopes)
    { # default scopes
        $myscopes=@()
        $myscopes+="User.Read.All"
    }
    $count_declines=0
    $connected_ok=$false
    do
    { # connect loop
        ###############
        Write-Host "Connect-MgGraph" -ForegroundColor Yellow
        # Get-MgContext -ErrorAction SilentlyContinue
        # Get-MgContext | Select-object -expand Scopes | Sort-Object # To show scopes
        $junk = Connect-MgGraph -scopes $myscopes -NoWelcome
        $domain_mg = Get-MgDomain -ErrorAction Ignore| Where-object IsDefault -eq $True | Select-object -ExpandProperty Id
        if ($domain_mg)
        { # has connection
            if ($domain)
            { # is specific domain
                if ($domain -eq $domain_mg)
                {
                    $connected_ok = $true
                }
                else
                { # no
                    Write-Host "Domains don't match. Using Disconnect-MgGraph and trying again."
                    PressEnterToContinue
                    $junk = Disconnect-MgGraph
                    #Continue #start loop again
                } # no
            } # is specific domain
            else
            { # didn't specify domain
                Write-Host "Already connected to Microsoft Graph.  (Use Disconnect-MgGraph to change domains)"
                Write-Host "Domain :" -NoNewline
                Write-Host $domain_mg -ForegroundColor Green
                $conn_choice = AskForChoice -Message "Use this connection? (No=Disconnect and retry)" -choices @("&Yes","&No","&Abort")
                If ($conn_choice -eq 0)
                { # yes, use connection
                    $connected_ok=$true
                }
                If ($conn_choice -eq 2)
                { # abort
                    $connected_ok=$false
                    Break # break out of loop
                }
                Else # 1: No
                { # don't use connection
                    $count_declines+=1 # keep track of how many no answers
                    $junk = Disconnect-MgGraph
                    if ($count_declines -gt 1)
                    { # max declines
                        If (-not (AskForChoice -Message "Keep Trying? (No=give up)"))
                        {
                            $connected_ok=$false
                            Break # break out of loop
                        }
                    } # max declines
                } # don't use connection
            } # didn't specify domain
        } # has connection
    } while (-not $connected_ok) # connect loop
    Return $connected_ok
}

Function GroupChildren ($DirectoryObjectId, $Recurse=$true)
{
    # Returns the children (users) of this group or role
    $myObjects = @()
    if ($Recurse)
    { # Recurse
        $Children = Get-MgGroupTransitiveMember -GroupId $DirectoryObjectId -ErrorAction Ignore
    }
    else
    { # # Non Recurse
        $Children = Get-MgGroupMember -GroupId $DirectoryObjectId -ErrorAction Ignore
    }
    ForEach ($Child in $Children)
    {
        $myObject = [PSCustomObject]@{
            Id  = $Child.Id
            Type   = $Child.AdditionalProperties.'@odata.type'.Replace("#microsoft.graph.","")
            displayName       = $Child.AdditionalProperties.displayName
            userPrincipalName = $Child.AdditionalProperties.userPrincipalName
            mail              = $Child.AdditionalProperties.mail
        }
        $MyObjects+=$myObject
    }
    Return $myObjects | Sort-Object mail
}
Function GroupParents ($DirectoryObjectId)
{
    # Returns the parents of this user (groups that this user is in)
    $myObjects = @()
    $ParentIDs = Get-MgDirectoryObjectMemberObject -DirectoryObjectId $DirectoryObjectId  -SecurityEnabledOnly:$False -ErrorAction Ignore
    ForEach ($ParentID in $ParentIDs)
    {
        $Parent = Get-MgDirectoryObject -DirectoryObjectId $ParentID
        $myObject = [PSCustomObject]@{
            Id  = $Parent.Id
            Type              = $Parent.AdditionalProperties.'@odata.type'.Replace("#microsoft.graph.","")
            displayName       = $Parent.AdditionalProperties.displayName
            GroupType         = $Parent.AdditionalProperties.groupTypes -join ", "
            securityEnabled   = $Parent.AdditionalProperties.securityEnabled
            mailEnabled       = $Parent.AdditionalProperties.mailEnabled
            mail              = $Parent.AdditionalProperties.mail
        }
        $MyObjects+=$myObject
    }
    Return $myObjects | Sort-Object mail
}

Function MgGroupCreate ($groupname)
{
    #MgGroupCreate - Creates a group in Azure/Entra (if necessary)
    #Usage: $strReturn,$group=MgGroupCreate "TestGroup"
    $strReturn = ""
    $MgGroup = @()
    $MgGroup += Get-MgGroup -All | Where-Object {$_.DisplayName -eq $groupname}
    If ($MgGroup) {$strReturn= "OK: found group id $($MgGroup[-1].Id)"}
    Else
    {
        $groupname_mail = $groupname.Replace(" ","_")
        $groupname_mail = $groupname_mail.Replace("(","_")
        $groupname_mail = $groupname_mail.Replace(")","_")
        $MgGroup = New-MgGroup -DisplayName $groupname -SecurityEnabled:$true -MailEnabled:$False -MailNickName $groupname_mail -ErrorAction Ignore
        Start-Sleep 1
        $MgGroup = @()
        $MgGroup += Get-MgGroup -All | Where-Object {$_.DisplayName -eq $groupname}
        If ($MgGroup) {$strReturn= "OK: created group $($MgGroup[-1].Id)"}
        Else {$strReturn= "ERR: couldn't create group $($groupname)"}
    }
    Return $strReturn,$MgGroup[-1]
}

########### End of File