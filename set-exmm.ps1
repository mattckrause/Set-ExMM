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
    
    .EXAMPLE
    Place an Exchange 2016 server into maintenance mode:
    set-exmm.ps1 -Server server1 -Action add

    Remove an Exchange 2016 server from maintenance mode:
    set-exmm.ps1 -Server server1 -action remove
    #>

#Read in parameters
param(
    [parameter(Position=0,Mandatory=$true,HelpMessage="Server Name")][string]$Server,
    [parameter(Position=1,Mandatory=$true,HelpMessage="add or remove")][string]$Action,
    [parameter(Position=2,Mandatory=$False,HelpMessage="Remote")][string]$loc
     )


if ($Action -eq "add") 
{
    write-host "$Server will be put into maintenance mode. Please wait..." -ForegroundColor Green

    write-host "`nSetting HubTransport to draining.." -ForegroundColor Blue
    set-servercomponentstate $Server -Component HubTransport -State Draining -Requester maintenance
        if($loc -eq "remote")
            {
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

    write-host "`nRestarting MSExchangeTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeTransport
    write-host "`nRestarting MSExchangeFrontEndTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeFrontEndTransport

    $moveTo = Read-Host "`n`nEnter the server name to redirect messages through (FQDN)"
    Write-Host "`nRedirecting messages from $Server to $moveTo..." -ForegroundColor Blue
    Redirect-Message -Server $Server -Target $moveTo -Confirm:$False

    $check = get-mailboxserver | fl DatabaseAvailabilityGroup
        if ($check -ne $null)
            {
                write-host "`n$Server is a DAG Member. Performing DAG maintenance mode proceedure..." -ForegroundColor Yellow
                write-host "`nSuspending cluster node..." -ForegroundColor Blue
                suspend-clusternode $Server

                write-host "`nMoving databases to another DAG member..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $true -Confirm:$False

                write-host "`nPreventing $Server from auto mounting databases..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Blocked -Confirm:$False
            }
        else
            {
                Write-Host "`n`n$Server is not a DAG member. Continuing...`n`n" -ForegroundColor Orange
            }

    Write-Host "`n`nPlacing $Server in maintenance mode..." -ForegroundColor Blue
    Set-ServerComponentState $Server -Component ServerWideOffline -State Inactive -Requester Maintenance

    Write-Host "`n`n$Server has been successfully put in maintenance mode.`n`n" -ForegroundColor Green

}
elseif ($Action -eq "remove") 
{
    write-host "$Server will be removed from maintenance mode. Please wait..." -ForegroundColor Green
    Set-ServerComponentState $Server -Component ServerWideOffline -State Active -Requester Maintenance

    $check = get-mailboxserver | fl DatabaseAvailabilityGroup
        if ($check -ne $null)
            {
                write-host "`n$Server is a DAG member. Performing maintenance mode removal proceedure for DAG member..." -ForegroundColor Yellow
                
                write-host "`nResuming cluster node..." -ForegroundColor Blue
                resume-clusternode $Server

                write-host "`nAllowing database copy activation..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $False -Confirm:$False

                write-host "`nEnabling DAG auto activation..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Unrestricted -Confirm:$False
            }
        else
            {
                Write-Host "`n`n$Server is not a DAG member. Continuing...`n`n" -ForegroundColor Orange
            }
    
    Write-Host "`n`nActivating HubTransport..." -ForegroundColor Blue
    Set-ServerComponentState $Server -Component HubTransport -State Active -Requester Maintenance
    
        if($loc -eq "remote")
            {
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

    write-host "`nRestarting MSExchangeTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeTransport
    write-host "`nRestarting MSExchangeFrontEndTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeFrontEndTransport
    
    Write-Host "`n`n$Server has been removed from maintenance mode.`n`n" -ForegroundColor Green
}   
else 
{
    write-host "Please inster a valid paramater for -Action." -ForegroundColor Red
}