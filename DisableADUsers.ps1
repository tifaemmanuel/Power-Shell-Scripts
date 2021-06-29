$inputFile = Import-CSV  "C:\admin\scripts\users to disable.csv"
foreach($line in $inputFile){
$name = $line.name 
Disable-ADAccount -identity $name
}