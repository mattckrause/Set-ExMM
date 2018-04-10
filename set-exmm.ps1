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
    [parameter(Position=1,Mandatory=$true,HelpMessage="add or remove")][string]$Action
     )

if ($Action -eq "add") 
{
    write-host "$Server will be put into maintenance mode. Please wait..." -ForegroundColor Green

    write-host "Setting HubTransport service to draining.." -ForegroundColor Blue
    set-servercomponentstate $Server -Component HubTransport -State Draining -Requester maintenance

    write-host "Restarting MSExchangeTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeTransport
    write-host "Restarting MSExchangeFrontEndTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeFrontEndTransport

    $moveTo = Read-Host "Enter a FQDN server name to Blueirect messages to"
    Write-Host "Blueirecting messages from $Server to $moveTo..." -ForegroundColor Blue
    Blueirect-Message -Server $Server -Target $moveTo -Confirm:$False

    $check = get-mailboxserver | fl DatabaseAvailabilityGroup
        if ($check -ne $null)
            {
                write-host "$Server is a DAG Member. Performing DAG Maintenance Mode proceedure..." -ForegroundColor Blue
                write-host "Suspending cluster node..." -ForegroundColor Blue
                suspend-clusternode $Server

                write-host "Moving databases to another DAG member..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $true -Confirm:$False

                write-host "Preventing $Server from auto mounting databases..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Blocked -Confirm:$False
            }
        else
            {
                Write-Host "$Server is not a DAG member. Continuing..." -ForegroundColor Orange
            }

    Write-Host "Placing $Server in maintenance mode..." -ForegroundColor Blue
    Set-ServerComponentState $Server -Component ServerWideOffline -State Inactive -Requester Maintenance

    Write-Host "$Server has been successfully put in maintenance mode. " -ForegroundColor Green

}
elseif ($Action -eq "remove") 
{
    write-host "$Server will be removed from maintenance mode. Please wait..." -ForegroundColor Green
    Set-ServerComponentState $Server -Component ServerWideOffline -State Active -Requester Maintenance

    $check = get-mailboxserver | fl DatabaseAvailabilityGroup
        if ($check -ne $null)
            {
                write-host "$Server is a DAG Member. Performing DAG maintenance mode removal proceedure..." -ForegroundColor Blue
                
                write-host "Resuming cluster node..." -ForegroundColor Blue
                resume-clusternode $Server

                write-host "Allowing database copy activation..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $False -Confirm:$False

                write-host "Enabling DAG auto activation..." -ForegroundColor Blue
                Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Unrestricted -Confirm:$False
            }
        else
            {
                Write-Host "$Server is not a DAG member. Continuing..." -ForegroundColor Orange
            }
    
    Write-Host "Activating HubTransport..." -ForegroundColor Blue
    Set-ServerComponentState $Server -Component HubTransport -State Active -Requester Maintenance
    
    write-host "Restarting MSExchangeTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeTransport
    write-host "Restarting MSExchangeFrontEndTransport service..." -ForegroundColor Blue
    Restart-Service MSExchangeFrontEndTransport
    
    Write-Host "$Server has been removed from maintenance mode." -ForegroundColor Green
}   
else 
{
    write-host "Please inster a valid paramater for -Action." -ForegroundColor Red
}