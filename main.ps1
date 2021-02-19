###Includes
. .\functions.ps1

# InitializeModules
$csvFile = Import-Csv $path

foreach ($line in $csvFile) {
  $upn = $line.UPN
  RemoveAccess($upn)
}



