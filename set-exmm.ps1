<#
    .DESCRIPTION
    set-exmm.ps1
    Matt Krause 2017

    This script will place the specified Exchange 2016 server into maintenance mode. It will also let you 
    remove the specified server from maintenance mode once the work has been completed.

    .PARAMETER Server
    The server that will be put in MM

    .PARAMETER Action
    The action the script will perform add/remove.

    -action add
    -action remove

    .PARAMETER Remote
    Is the script being run from a remote machine or not.

    -Remote
    
    .EXAMPLE
    Place an Exchange 2016 server into maintenance mode:
    set-exmm.ps1 -Server server1 -Action add

    Run from a remote location
    set-exmm.ps1 -Server server1 -Action add -Remote

    Remove an Exchange 2016 server from maintenance mode:
    set-exmm.ps1 -Server server1 -action remove
    #>

#Read in parameters
param(
    [parameter(Position=0,Mandatory=$true,HelpMessage="Server Name")][string]$Server,
    [parameter(Position=1,Mandatory=$true,HelpMessage="add or remove")][string]$Action,
    [parameter(Position=2,Mandatory=$False,HelpMessage="Remote")][switch]$Remote
     )

#check if we are adding or removing a server from MM
if ($Action -eq "add") #Add Exchange server to MM.
{
    #Set Hubtransport to draining
    write-host "$Server will be put into maintenance mode. Please wait..." -ForegroundColor Green
    write-host "`nSetting HubTransport to draining.." -ForegroundColor Blue
    set-servercomponentstate $Server -Component HubTransport -State Draining -Requester maintenance
    #Check if the script is being ran remotely or not
    if($Remote -eq $True)
        {
            #Restart services on a remote server
            $MSET = get-service -ComputerName $Server -Name MSExchangeTransport
            $MSEfeT = get-service -ComputerName $Server -Name MSExchangeFrontEndTransport
            Write-Host "`nRestarting MSExchangeTransport service on $Server..."
            Restart-Service -InputObject $MSET -Verbose
            $MSET.Refresh()
            $MSET
            Write-Host "`nRestarting MSExchangeFrontEndTransport service on $Server..."
            Restart-Service -InputObject $MSEfeT -Verbose
            $MSEfeT.Refresh()
            $MSEfeT
            #Get the computer to redirect messages to. Must be FQDN.
            $moveTo = Read-Host "`n`nEnter the server name to redirect messages through (FQDN)"
            Write-Host "`nRedirecting messages from $Server to $moveTo..." -ForegroundColor Blue
            Redirect-Message -Server $Server -Target $moveTo -Confirm:$False
            #Check if the server is a member of a DAG.
            $check = get-mailboxserver $Server | format-list DatabaseAvailabilityGroup
            if ($check.DatabaseAvailabilityGroup -ne $null)
                {
                    #Suspend DAG cluster node of a remote Exchange Server.
                    write-host "`n$Server is a DAG Member. Performing DAG maintenance mode proceedure..." -ForegroundColor Yellow
                    write-host "`nSuspending cluster node..." -ForegroundColor Blue
                    Invoke-Command -ComputerName $Server -ScriptBlock {param($SRV) suspend-clusternode $SRV} -ArgumentList $Server
                    #Move all mounted databases to a different DAG member.
                    write-host "`nMoving databases to another DAG member..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $true -Confirm:$False
                    #Disable databases from auto mounting on the server.
                    write-host "`nPreventing $Server from auto mounting databases..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Blocked -Confirm:$False
                }
            else
                {
                    Write-Host "`n`n$Server is not a DAG member. Continuing...`n`n" -ForegroundColor Cyan
                }
        }
    else
        {
            #Restart Transport services on local machine
            write-host "`nRestarting MSExchangeTransport service..." -ForegroundColor Blue
            Restart-Service MSExchangeTransport
            write-host "`nRestarting MSExchangeFrontEndTransport service..." -ForegroundColor Blue
            Restart-Service MSExchangeFrontEndTransport
            #Provide server to redirect messages to.
            $moveTo = Read-Host "`n`nEnter the server name to redirect messages through (FQDN)"
            Write-Host "`nRedirecting messages from $Server to $moveTo..." -ForegroundColor Blue
            Redirect-Message -Server $Server -Target $moveTo -Confirm:$False
            #Check if the server is a DAG member.
            $check = get-mailboxserver | format-list DatabaseAvailabilityGroup
            if ($check.DatabaseAvailabilityGroup -ne $null)
                {
                    #Suspend DAG cluster node on local Exchange server.
                    write-host "`n$Server is a DAG Member. Performing DAG maintenance mode proceedure..." -ForegroundColor Yellow
                    write-host "`nSuspending cluster node..." -ForegroundColor Blue
                    suspend-clusternode $Server
                    #Move all mounted databases to a different DAG member.
                    write-host "`nMoving databases to another DAG member..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $true -Confirm:$False
                    #Disable databases from auto mounting on the server
                    write-host "`nPreventing $Server from auto mounting databases..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Blocked -Confirm:$False
                }
            else
                {
                    Write-Host "`n`n$Server is not a DAG member. Continuing...`n`n" -ForegroundColor Cyan
                }
        }
    #Complete putting the Exchange server in MM.
    Write-Host "`n`nPlacing $Server in maintenance mode..." -ForegroundColor Blue
    Set-ServerComponentState $Server -Component ServerWideOffline -State Inactive -Requester Maintenance
    Write-Host "`n`n$Server has been successfully put in maintenance mode.`n`n" -ForegroundColor Green
}
elseif ($Action -eq "remove") #Remove Exchange server from MM.
{
    #Take Exchange server out of MM.
    write-host "$Server will be removed from maintenance mode. Please wait..." -ForegroundColor Green
    Set-ServerComponentState $Server -Component ServerWideOffline -State Active -Requester Maintenance
    #Check if the script is being ran remotely or not.
    if($Remote -eq $True)
        {
            #Check if the Exchange server is a DAG member.
            $check = get-mailboxserver $Server | format-list DatabaseAvailabilityGroup
            if ($check.DatabaseAvailabilityGroup -ne $null)
                {
                    #Resume the cluster node remotely.
                    write-host "`n$Server is a DAG member. Performing maintenance mode removal proceedure for DAG member..." -ForegroundColor Yellow
                    write-host "`nResuming cluster node..." -ForegroundColor Blue
                    Invoke-Command -ComputerName $Server -ScriptBlock {param($SRV) Resume-ClusterNode $SRV} -ArgumentList $Server
                    #Allow Exchange server to host databases again.
                    write-host "`nAllowing database copy activation..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $False -Confirm:$False
                    #Allow Exchange to auto mount databases on the server.
                    write-host "`nEnabling DAG auto activation..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Unrestricted -Confirm:$False
                }
            else
                {
                    Write-Host "`n`n$Server is not a DAG member. Continuing...`n`n" -ForegroundColor Cyan
                }
            #Enable HubTranport functionality.
            Write-Host "`n`nActivating HubTransport..." -ForegroundColor Blue
            Set-ServerComponentState $Server -Component HubTransport -State Active -Requester Maintenance
            #Restart Transpor services of remote server.
            $MSET = get-service -ComputerName $Server -Name MSExchangeTransport
            $MSEfeT = get-service -ComputerName $Server -Name MSExchangeFrontEndTransport
            Write-Host "`nRestarting MSExchangeTransport service on $Server..."
            Restart-Service -InputObject $MSET -Verbose
            $MSET.Refresh()
            $MSET
 
            Write-Host "`nRestarting MSExchangeFrontEndTransport service on $Server..."
            Restart-Service -InputObject $MSEfeT -Verbose
            $MSEfeT.Refresh()
            $MSEfeT
        }
    else #Script not run remotely.
        {
            #Check if the Exchange server is a DAG member.
            $check = get-mailboxserver $Server | format-list DatabaseAvailabilityGroup
            if ($check.DatabaseAvailabilityGroup -ne $null) #DAG Member
                {
                    #Resume DAG cluster node on local server.
                    write-host "`n$Server is a DAG member. Performing maintenance mode removal proceedure for DAG member..." -ForegroundColor Yellow
                    write-host "`nResuming cluster node..." -ForegroundColor Blue
                    resume-clusternode $Server
                    #Allow databases to be mounted on this server.
                    write-host "`nAllowing database copy activation..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $False -Confirm:$False
                    #Allow Exchange to auto mount databases on this server.
                    write-host "`nEnabling DAG auto activation..." -ForegroundColor Blue
                    Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Unrestricted -Confirm:$False
                }
            else
                {
                    Write-Host "`n`n$Server is not a DAG member. Continuing...`n`n" -ForegroundColor Cyan
                }
            #Enable HubTransport functionality.
            Write-Host "`n`nActivating HubTransport..." -ForegroundColor Blue
            Set-ServerComponentState $Server -Component HubTransport -State Active -Requester Maintenance
            #Restart transport services on the local server.
            write-host "`nRestarting MSExchangeTransport service..." -ForegroundColor Blue
            Restart-Service MSExchangeTransport
            write-host "`nRestarting MSExchangeFrontEndTransport service..." -ForegroundColor Blue
            Restart-Service MSExchangeFrontEndTransport
        }
    Write-Host "`n`n$Server has been removed from maintenance mode.`n`n" -ForegroundColor Green
}   
else 
{
    write-host "Please inster a valid paramater for -Action." -ForegroundColor Red
}