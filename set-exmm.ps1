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
    write-host "$Server will be put into maintenance mode. Please wait..."    
}
elseif ($Action -eq "remove") 
{
    write-host "$Server will be removed from maintenance mode. Please wait..."    
}
else 
{
    write-host "Please inster a valid paramater for -Action."    
}