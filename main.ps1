###Includes
. .\functions.ps1

InitializeModules
$csvFile = Import-Csv $path

foreach ($line in $csvFile) {
  $upn = $line.User
  RemoveAccess($upn)
}



