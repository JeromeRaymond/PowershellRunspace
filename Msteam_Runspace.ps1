#Initialisation des variables
$csv=$null
$Results=$null

# suppression de runspace "fantomes"
$PreviousRS = get-runspace | where-object {($_.id -ne 1)}
$PreviousRS.dispose()

# detection de machines à tester
import-module activedirectory
$searchdate = (get-date).AddDays(-30) 
$ComputersAD= Get-ADcomputer -Filter * -properties * | where-object {$_.LastLogonDate -ge $searchdate} | select CN
$computers = $computersAD | foreach ($_.CN){write $_.CN}

# définition du Pool (creation des slots)
$pool = [RunspaceFactory]::CreateRunspacePool(1,10)
$pool.ApartmentState = "MTA"
$pool.Open()
$runspaces = @()

# définition des commandes à executer dans les jobs
$scriptblock = {
Param (
[string]$server
)
Invoke-Command -ComputerName $server -ScriptBlock {Get-NetLbfoTeam | select-object PSComputername,Name,Status,Teamingmode,Members}
}

# creation des jobs par machine et lancement des jobs
foreach ($ComputerName in $Computers)
{
$runspace = [PowerShell]::Create()
$null = $runspace.AddScript($scriptblock)
$null = $runspace.AddArgument($ComputerName)
$runspace.RunspacePool = $pool
$runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
}

# affiche toutes les secondes l'utilisations des slots + statistiques
while ($runspaces.Status -ne $null)
{
start-sleep 1
cls
get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $false)  -and ($_.runspaceAvailability -like "Available")}
get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $true)}
$slt_encours = get-runspace | where-object {($_.id -ne 1) -and ( $_.runspaceisremote -eq $true)}
$slt_tot = get-runspace | where-object {($_.id -ne 1) -and ($_.runspaceisremote -eq $false)}

write-host "Nbre Objets total= " $runspaces.count
write-host  "Nbre slots totaux= " $slt_tot.count
write-host  "Nbre slots utilisés= " $slt_encours.count
write-host  "Nbre objets restants =" $runspaces.Status.IsCompleted.count

$completed = $runspaces | Where-Object { $_.Status.IsCompleted -eq $true }

foreach ($runspace in $completed)
{
$Results += $runspace.Pipe.EndInvoke($runspace.Status)
$runspace.Status = $null
}

}

# Export de tous les logs dans un fichier commun
$Results | select-object PSComputername,Name,Status,Teamingmode,Members | export-csv -NoTypeInformation -path "c:\temp\msteam.csv"

# Fermeture les connexions et suppression des slots du Pool
$pool.Close()
$pool.Dispose()
