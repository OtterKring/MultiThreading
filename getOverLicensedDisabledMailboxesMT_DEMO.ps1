[CmdletBinding()]
param (
    [Parameter()]
    [byte]
    $Threads = 10
)

#region SINGLETHREADED_EXECUTION


#region PREPARE_CONNECTIONDATA

Write-Host 'Preparing MgGraph connection data'
# Get MSGraph App-Registration data from Vault and build the splatting concurrent dictionary
$MgMetadata = ( Get-SecretInfo -Name GraphInfo ).Metadata
$splat_MgGraph = [System.Collections.Concurrent.ConcurrentDictionary[string,psobject]]::new()
foreach ( $key in $MgMetadata.Keys) { $splat_MgGraph.$key = $MgMetadata.$key }
$splat_MgGraph.NoWelcome = $true
$splat_MgGraph.ErrorAction = 'Stop'

$splat_MgGModule = [System.Collections.Concurrent.ConcurrentDictionary[string,psobject]]::new()
$splat_MgGModule.Name = 'Microsoft.Graph.Users'
$splat_MgGModule.Cmdlet = 'Get-MgUser','Get-MgSubscribedSku'
$splat_MgGModule.ErrorAction = 'Stop'

Write-Host 'Preparing ExchangeOnline connection data'
# Get ExchangeOnline App-Registration data from Vault and build the splatting concurrent dictionary
$MgMetadata = ( Get-SecretInfo -Name EXOInfo ).Metadata
$splat_EXO = [System.Collections.Concurrent.ConcurrentDictionary[string,psobject]]::new()
foreach ( $key in $MgMetadata.Keys) { $splat_EXO.$key = $MgMetadata.$key }
$splat_EXO.ShowBanner = $false
$splat_EXO.ErrorAction = 'Stop'
$splat_EXO.CommandName = 'Get-Mailbox'

$splat_EXOModule = [System.Collections.Concurrent.ConcurrentDictionary[string,psobject]]::new()
$splat_EXOModule.Name = 'ExchangeOnlineManagement'
$splat_EXOModule.ErrorAction = 'Stop'

# get rid of the unnecessary secret data
Remove-Variable -Name MgMetadata

#endregion PREPARE_CONNECTIONDATA

if ( -not ( Get-MgUser -Top 1 -ErrorAction SilentlyContinue ) ) {
    Import-Module @splat_MgGModule
    Connect-MgGraph @splat_MgGraph
}
if ( -not ( Get-EXOMailbox -ResultSize 1 -ErrorAction SilentlyContinue ) ) {
    Import-Module @splat_EXOModule
    Connect-ExchangeOnline @splat_EXO
}

#region GET_LICENSES

# create a skuid->skupartnumber Concurrent Dictionary (for multi-threaded access) for quick license lookup
Write-Host 'Caching Exchange licenses ... ' -NoNewLine
$Skus = Get-MgSubscribedSku
$MgSubscribedSku = [System.Collections.Concurrent.ConcurrentDictionary[string,string]]::new()
foreach ( $sku in $Skus ) {
    if ( $sku.ServicePlans.ServicePlanName -contains 'Exchange_S_Deskless' -or $sku.ServicePlans.ServicePlanName -contains 'Exchange_S_Enterprise' ) {
        $MgSubscribedSku.($sku.SkuId) = $sku.SkuPartNumber
    }
}
Remove-Variable -Name Skus
Write-Host "$($MgSubscribedSku.Keys.Count) found" -ForegroundColor Green

#endregion GET_LICENSES


#region GET_MAILBOXES

# get all disabled mailboxes and make a Concurrent Dictionary (for multi-threaded access) based on the MgGraph Id, delivering the MgGraphId itself and if an archive exists or not
Write-Host 'Querying disabled mailboxes ... ' -NoNewLine
$DisabledMailboxes = Get-ExoMailbox -Filter "ExchangeUserAccountControl -eq 'AccountDisabled' -and ArchiveState -eq 'None'" -ResultSize unlimited
$DisabledMailbox = [System.Collections.Concurrent.ConcurrentDictionary[string,hashtable]]::new()
foreach ( $dmbx in $DisabledMailboxes ) {
    $DisabledMailbox.($dmbx.ExternalDirectoryObjectId) = @{
        Id = $dmbx.ExternalDirectoryObjectId
        Mailbox = $dmbx
    }
}
Remove-Variable -Name DisabledMailboxes
Write-Host "$($DisabledMailbox.Keys.Count) found" -ForegroundColor Green

#endregion GET_MAILBOXES


#region GET_MGUSERS

# get all disabled MgGraph user within onprem OU OU=Deactivated Users,DC=agrana,DC=net with their licenses, country and usagelocation
Write-Host 'Querying potentially overlicensed disabled mggraph users with mailboxes ... ' -NoNewLine
$MgProperties = 'Id','UserPrincipalName','Displayname','AssignedLicenses','OnpremisesDistinguishedName','Country','UsageLocation'
[System.Collections.Concurrent.ConcurrentQueue[psobject]]$MgDeactivatedUsers = Get-MgUser -Filter "AccountEnabled eq false" -All <# -Top 100 #> -Property $MgProperties -ErrorAction SilentlyContinue |
    Where-Object {
        process {
            $_.OnPremisesDistinguishedName -like '*,OU=Deactivated Users,DC=agrana,DC=net' -and
            $_.AssignedLicenses.SkuId.Where{ $_ -in $MgSubscribedSku.Keys }
         }
    } |
    Select-Object $MgProperties
Write-Host "$($MgDeactivatedUsers.Count) found" -ForegroundColor Green

#endregion GET_MGUSERS

#endregion SINGLETHREADED_EXECUTION



#region MULTITHREADED_EXECUTION

# $oldProgressPreference = $ProgressPreference
$ProgressPreference = 'Continue'


1..$Threads | Foreach-Object -ThrottleLimit $Threads -Parallel {

    #region MTVariables

    $MTsplat_MgGraph = $using:splat_MgGraph
    $MTsplat_MgGModule = $using:splat_MgGModule
    $MTsplat_EXO = $using:splat_EXO
    $MTsplat_EXOModule = $using:splat_EXOModule
    $MTMgSubcribedSku = $using:MgSubscribedSku
    $MTDisabledMailbox = $using:DisabledMailbox
    $MTMgDeactivatedUsers = $using:MgDeactivatedUsers

    $MgUsercount = $MTMgDeactivatedUsers.Count
    $ThreadsWidth = ($using:Threads).ToString().Length

    #endregion MTVariables

    #region Build_Namespace_Connections

    if ( -not ( Get-Command -Name Get-MgUser -ErrorAction SilentlyContinue ) ) {
        Import-Module @MTsplat_MgGModule
        Connect-MgGraph @MTsplat_MgGraph
    }
    if ( -not ( Get-Command -Name Get-ExoMailboxStatistics -ErrorAction SilentlyContinue ) ) {
        Import-Module @MTsplat_EXOModule
        Connect-ExchangeOnline @MTsplat_EXO
    }

    #endregion Build_Namespace_Connections

    #region MTFunctions

    function returnItem ( $InputObject ) {
        [PSCustomObject]@{
            Id = $InputObject.Id
            DisplayName = $InputObject.DisplayName
            AssignedLicenses =  foreach ( $skuid in $InputObject.AssignedLicenses.SkuId ) { [PSCustomObject]@{ SkuId = $skuid; SkuPartNumber = $MTMgSubcribedSku.$skuid } }
        }
    }

    #endregion MTFunctions

    #region MT_Workerloop

    while ( $MTMgDeactivatedUsers.Count -gt 0 ) {

        $item = $null

        if ( $MTMgDeactivatedUsers.TryDequeue( [ref]$item ) ) {

            Write-Progress -Id 0 -Activity 'Filtering objects' -Status ( "Thread: {0:$('0' * $ThreadsWidth)}, Account: {1}" -f $_, $item.DisplayName ) -PercentComplete ( ( $MgUsercount - $MTMgDeactivatedUsers.Count ) * 100 / $MgUsercount )

            if ( $MTDisabledMailbox.($item.Id) ) {

                $MgMbxUserPurpose = ( Get-MgUser -UserId $item.Id -Property MailboxSettings -ErrorAction SilentlyContinue ).MailboxSettings.UserPurpose
                if ( $MgMbxUserPurpose ) {
                    $mbxSize = ( Get-ExoMailboxStatistics $item.Id -PropertySets Minimum -ErrorAction SilentlyContinue ).TotalItemSize
                    if ( $mbxSize ) {
                        $mbxSize = $mbxSize.ToString().Split('(')[1] -replace '\D+'     # Solution from Internet (~40% slower): .Split("(")[1].Split(" ")[0].Replace(",","")
                        if ( 50GB -gt $mbxSize ) {
                            returnItem $item
                        }
                    }
                }
            }
        }
    }

    #endregion MT_Workerloop

    Write-Progress -Id 0 -Activity 'Filtering objects' -Completed

    # $ProgressPreference = $oldProgressPreference

}

#endregion MULTITHREADED_EXECUTION