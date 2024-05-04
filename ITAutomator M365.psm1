
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
2024-05-03 Initial Version
MgGroupCreate
O365Connect -ManualMode $true (for Connect-ExchangeOnline)

Alphabetical list of functions
MgGroupCreate - Creates a group in Azure/Entra (if neceeesary). This function assumes Connect-MgGraph has already been called
#>

Function MgGroupCreate ($groupname)
{
    #MgGroupCreate - Creates a group in Azure/Entra (if necessary). This function assumes Connect-MgGraph has already been called
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

Function O365Connect
{
    Param (
         [string] $scriptXML
        ,[String] $Domain =""
        ,[boolean] $ManualMode=$false
    )
    # $O365_PasswordXML   = $scriptDir+ "\O365_Password.xml"
    # Write-Host "XML: $($O365_PasswordXML)"
    # ## ----------Connect/Save Password
    # $PSCred=O365Connect($O365_PasswordXML)
    #
    #
    if ($ManualMode)
    {
        if (!(Get-Command "Connect-ExchangeOnline" -errorAction SilentlyContinue))
        {
            Write-host "WARNING: Your powershell environment doesn't have the command Connect-ExchangeOnline."
            Write-Host "Install these from a Powershell (as admin) prompt:"
            Write-Host "------------------------"
            Write-Host "Install-Module -Name ExchangeOnlineManagement"
		    Pause
            exit
        }
        ##
        $connected = $false
        $conns = @(Get-ConnectionInformation)
        if ($conns.count -gt 0)
            { # has conns
                $curr_conn=(Get-ConnectionInformation)[-1] # last connection
                if (($curr_conn.State = "Connected") -and ($curr_conn.UserPrincipalName -match $domain))
                {
                    $connected = $true
                }
            } # has conns
        if (!($connected))
        { # not connected
            Try
            {
                Connect-ExchangeOnline -ShowBanner:$false
            }
            Catch
            {
                Write-Host "------------------------"
                Write-Host "Connect-ExchangeOnline was canceled or didn't complete."
                Write-Host "------------------------"
                Pause
                exit
            }
        } # not connected
        ##
        $PSCred=[PSCustomObject]@{
            UserName = $Domain
            Password = "****"
            }
        Return $PSCred
    }

    if (!(Get-Command "Connect-MSOLService" -errorAction SilentlyContinue))
    {
        Write-host "WARNING: Your powershell environment doesn't have the command Connect-MSOLService."
        Write-Host "Install these from a Powershell (as admin) prompt:"
        Write-Host "------------------------"
        Write-Host "Install-Module AzureAD"
		Write-Host "Install-Module MSOnline"
		Pause
        exit
    }

    ## Globals Init with defaults                                            
    $O365Globals = @{}
    $O365Globals.Add("CreateDate",(Get-Date).ToString("yyyy-MM-dd hh:mm:ss"))
    $O365Globals.Add("LastUsedDate",(Get-Date).ToString("yyyy-MM-dd hh:mm:ss")+" ["+${env:COMPUTERNAME}+" "+${env:USERNAME}+"]")
    ## Globals Load from XML                                                 
    $O365Globals=GlobalsLoad $O365Globals $scriptXML $false

    #if ($passes.Count -eq 1)     {        $admin_creds = $passes[0]    }
    #else     {        ### pick one        $admin_creds = $passes[0]    }
    ## 

    ###
    $done=$true
    Do
    { ######### Choice loop

        #### Gets a list of eligibal passwords for this user/pc combination
        $passes = @($O365Globals.Passwords | Where-Object {($_.hostname -eq ${env:COMPUTERNAME}) -and ($_.username -eq ${env:USERNAME})}| Sort-Object adminuser)
        ###
        $choice_list = @() # $null
        $choice_default = ""
        $i = 1
        Write-Host "-----------------"
        Write-Host "Select an account (last used $($O365Globals.LastUsedDate))"
        Write-Host "-----------------"
        ForEach ($pass in $passes)
        {
            $choice_obj=@(
                [pscustomobject][ordered]@{
                number=$i
                adminuser=$pass.adminuser
                adminpass=$pass.adminpass
                }
            )
            $choice_descrip = ""
            if ($domain -ne "")
            { #domain match requested
                if ($pass.adminuser.Split("@")[1].ToLower() -eq $Domain.ToLower())
                { #match
                    $choice_default = $i
                    $choice_descrip = " [match for '$($Domain)']"
                } #match
            } #domain match requested
            Write-Host "  $($i)]  $($choice_obj.adminuser) $($choice_descrip)"
            ### append object
            $i+=1
            $choice_list +=$choice_obj
        }
        $i-=1
        Write-Host "-----------------"
        #### Get input
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
		if ($i -eq 0)
			{$msg = "Enter a username`r`n[(blank) to Cancel]"
            Write-Host $msg
            $choice = [Microsoft.VisualBasic.Interaction]::InputBox($msg, "User", $choice_default)
            }
		else
			{
            $msg = "Enter a number (1-$($i)) <OR> a username`r`n[Xn to delete entry (n), (blank) to Cancel]"
            Write-Host $msg
            $menu=@()
            if ($choice_default)
            {
                $choice_obj=@(
                [pscustomobject][ordered]@{
                number=$choice_default
                adminuser="<Default choice $($Domain)>"
                })
                $menu += $choice_obj
            }
            $choice_obj=@(
                [pscustomobject][ordered]@{
                number="<Add>"
                adminuser="<Add a new entry>"
                })
            $menu += $choice_obj
            $choice_obj=@(
                [pscustomobject][ordered]@{
                number="<Delete>"
                adminuser="<Delete an entry>"
                })
            $menu += $choice_obj
            $menu += $choice_list |Select-Object number,adminuser
            $Result = $menu | Out-GridView -PassThru  -Title 'Choose Org'
            $choice = $Result.number
            ## if X , relist for delete
            if ($choice -eq "<Delete>")
            {
                $menu = $choice_list |Select-Object number,adminuser
                $Result = $menu | Out-GridView -PassThru  -Title 'Choose Org to <Delete>'
                $choice = "X" + $Result.number
                if (-not($choice)) {write-host "Aborted by user.";return $null}
            }

            ## if Add, prompt for admin
            if ($choice -eq "<Add>")
                {
                    $msg = "Enter a username`r`n[(blank) to Cancel]"
                    Write-Host $msg
                    $choice = [Microsoft.VisualBasic.Interaction]::InputBox($msg, "User", $choice_default)
                    if (-not($choice)) {write-host "Aborted by user.";return $null}
                }
            }
        
        ###
        $O365Globals.LastUsedDate = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss")+" ["+${env:COMPUTERNAME}+" "+${env:USERNAME}+"]"
        if (-not($choice)) {write-host "Aborted by user.";return $null}
        if ($choice -match '^\d+$')
        { ### picked a number to use
            $pass = @($choice_list | Where-Object {($_.number -eq $choice)})
            if (!($pass))
            { ## invalid number
                Write-Host "[INVALID CHOICE]"
                $done=$false
            }
            else
            { ## try the selected password
                Try
                {
                    ### ConvertTo-SecureString: SecureString_Plaintext >> SecureString (PSCreds use SecureString, XML stores SecureString_Plaintext)
                    $pass_secstr = ConvertTo-SecureString $pass.adminpass -ErrorAction Stop
                    $PSCred=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $pass.adminuser , $pass_secstr -ErrorAction Stop
                    # -------------------------------------------------------
                    # Create creds based on saved password using DPAPI.  
                    # DPAPI is Microsoft's data protection method to store passwords at rest.  The files are only decryptable on the machine / user that created them.
                    # 
                    # Decrypt methods below are OK for debugging, as long as the decrypted values aren't saved
                    #
                    # Decrypt method 1
                    # [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($pass_secstr_plaintext))
                    #
                    # Decrypt method 2
                    # $PSCred.GetNetworkCredential().password
                    # -------------------------------------------------------
                    # Try to connect using these credentials
                    Connect-MSOLService -Credential $PSCred -ErrorAction Stop
                    $done=$true
                    ## Globals Persist to XML
                    GlobalsSave $O365Globals $scriptXML
                }
                Catch
                {
                    Write-Warning "[Invalid or no password value for [$($pass.adminuser)]. Perhaps delete it and try again."
                    $done=$false
                }
            } ## try the selected password
        } ### picked a number to use
        elseif ($choice.Substring(0,1).tolower() -eq "x")
        { ### picked a number to delete
            $choice=$choice.substring(1)
            $pass = @($choice_list | Where-Object {($_.number -eq $choice)})
            if (!($pass))
            { ## invalid delete number
                Write-Host "[INVALID DELETE CHOICE]"
                $done=$false
            } ## invalid delete number
            else
            { ## valid delete number
                Write-Host "[DELETING $($choice): $($pass.adminuser)]"
                # Return everything BUT an exact match (deletes old password if it exists)
                $new_passes = @($O365Globals.Passwords | Where-Object {-not (($_.hostname -eq ${env:COMPUTERNAME}) -and ($_.username -eq ${env:USERNAME}) -and ($_.adminuser -eq $pass.adminuser))})
                $O365Globals.Passwords = $new_passes
                #
                $done=$false
                ## Globals Persist to XML
                GlobalsSave $O365Globals $scriptXML
            } ## valid delete number
        } ### picked a number to delete
        else
        { ### New adminuser
            $PSCred = Get-Credential -Message "Enter O365 Admin Password" -UserName $choice
            if (!($PSCred))
            { # no creds
                Write-Host "No creds entered"
                $done=$false
            } # no creds
            else
            { # got creds
                Write-Host "Trying to connect to MS Online using $($PSCred.UserName)"
                Try
                {
                    # -------------------------------------------------------
                    # Try to connect using these credentials
                    Connect-MSOLService -Credential $PSCred -ErrorAction Stop
                    $done=$true
                }
                Catch
                {
                    Write-Warning "[Invalid or no password value for [$($PSCred.UserName)]. Perhaps delete it and try again."
                    $done=$false
                }
                if ($done)
                    { # they worked
                    #### Save globals
                    ## Create a new object based on user entries
                    $obj = [pscustomobject]@{                       
                            adminuser = $PSCred.UserName
                            adminpass = ($PSCred.Password | ConvertFrom-SecureString)
                            hostname  = ${env:COMPUTERNAME}
                            username  = ${env:USERNAME}
                            }
                    if ($O365Globals.Passwords)
                    { # old passwords found
                        # Return everything BUT an exact match (deletes old password if it exists)
                        $new_passes = @($O365Globals.Passwords | Where-Object {-not (($_.hostname -eq ${env:COMPUTERNAME}) -and ($_.username -eq ${env:USERNAME}) -and ($_.adminuser -eq $obj.adminuser)) })
                        $O365Globals.Passwords = $new_passes
                        #
                        #add this password
                        $O365Globals.Passwords+=$obj
                    } # old passwords found
                    else
                    { # no old passwords found
                        $O365Globals.Passwords=@()
                        $O365Globals.Passwords+=$obj
                    } # no old passwords found
                    ## Globals Persist to XML
                    GlobalsSave $O365Globals $scriptXML
                    } # they worked
            } # got creds
        } ### New adminuser
    } Until ($done) ######### Choice loop
    # Return credentials
    $PSCred
}

Function ConnectMgGraph ($myscopes=$null, $domain=$null)
{
    if (-not $myscopes)
    {
        $myscopes=@()
        $myscopes+="User.Read.All"
    }
    do # connect loop
    {
        $connectedall_ok = $true
        ###############
        Write-Host "Connect-MgGraph" -ForegroundColor Yellow
        #Read-Host "Press (Enter) to continue"
        # Load the module and show results
        $module= "Microsoft.Graph.Authentication" ; Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
        #Read-Host "Press (Enter) to continue"
        $connected_ok=$false
        #Get-MgContext -ErrorAction SilentlyContinue
        $domain_mg = Get-MgDomain -ErrorAction Ignore| Where-object IsDefault -eq $True | Select-object -ExpandProperty Id
        if ($domain_mg)
        { # has Get-MgContext
            # Get-MgContext | Select-object -expand Scopes | Sort-Object # To show scopes
            #$domain_mg = Get-MgDomain | Where-object IsDefault -eq $True | Select-object -ExpandProperty Id
            if ($domain)
            { # we need it to be this domain
                if ($domain -ne $domain_mg)
                {
                    Write-Host "Domains don't match. Using Disconnect-MgGraph and trying again."
                    Read-Host "Press (Enter) to continue"
                    Disconnect-MgGraph
                    Continue #start loop again
                }
            } # we need it to be this domain
            else
            { # didn't specify domain
                Write-Host "Already connected to Microsoft Graph.  (Use Disconnect-MgGraph to change domains)"
                Write-Host "Domain :" -NoNewline
                Write-Host $domain_mg -ForegroundColor Green
                If (AskForChoice -Message "Use this connection? (No=Disconnect and reconnect)")
                {
                    $connected_ok=$true
                }
                Else
                {
                    Disconnect-MgGraph
                    Continue #Start for loop again
                }
            } # didn't specify domain
        } # has Get-MgContext
        if (-not $connected_ok)
        {
            Connect-MgGraph -scopes $myscopes -NoWelcome
        }
        $domain = Get-MgDomain | Where-object IsDefault -eq $True | Select-object -ExpandProperty Id
        Write-Host "Domain :" -NoNewline
        Write-Host $domain -ForegroundColor Green
        if ($connectedall_ok)
        {Break}
    } while ($true) # connect loop
    Return $connectedall_ok
}

Function ConnectExchangeOnline ($domain=$null)
{
    do # connect loop
    {
        $connectedall_ok = $true
        ###############
        Write-Host "Connect-ExchangeOnline" -ForegroundColor Yellow
        #Read-Host "Press (Enter) to continue"
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
                Read-Host "Press (Enter) to continue"
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
########### End of File