
################################
#
#  Code de récupération de la volumétrie et de l'état des disques
#
#################################


$error="SilentlyContinue"
$date = Get-Date -f g

# récupéréation de la liste des PC à checker
import-module activedirectory
$searchdate = (get-date).AddDays(-30) 
$Computer= Get-ADcomputer -Filter * -properties * | where-object {$_.LastLogonDate -ge $searchdate} |select name

# code à executer
$teamresult = $computer | ForEach {Invoke-Command -ComputerName $_.name -ScriptBlock {Get-NetLbfoTeam | select-object PSComputername,Name,Status,Teamingmode,Members}}

# Logs
$teamresult | select-object PSComputername,Name,Status,Teamingmode,Members | export-csv -NoTypeInformation -path "c:\temp\msteam.csv"
