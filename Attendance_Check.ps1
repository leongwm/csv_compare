#Date formatting
$date = Get-Date -UFormat "%d %b %Y"
$date2 = Get-Date -UFormat "%d/%m/%Y"

#Extract CSV zip
Expand-Archive -Path .\"Attendance Form USM Karate ($date).csv.zip" -DestinationPath .\ 
echo "============================================================="
echo "DO NOT CLOSE THIS WINDOW. IT WILL CLOSE AUTOMATICALLY"
echo "============================================================="

echo "Extracting Zip File... "
echo ""

Start-Sleep -s 5

#Get matrix number from student list
$csv1 = (Import-Csv -Path .\Attendance_full.csv)."Matrix Number:"
echo "Getting Matrix Number from Full Student List..."

#Get matrix number from attendance list
$csv2 = (Import-Csv -Path .\"Attendance Form USM Karate ($date).csv")."Matrix Number:"
echo "Getting Matrix Number from Today's Attendance List..."

 
#Comparing duplicate values from student list and attendance list
$removedDuplicateRows = Import-Csv -Path .\Attendance_full.csv | Where-Object {$_."Matrix Number:" -notin $csv2}
$removedDuplicateRows | Export-Csv -Path .\Absent_Output_$date.csv -NoTypeInformation

echo ""
echo "Comparing results..."

Start-Sleep -s 2

#Count total number of students
$total = Import-Csv .\Attendance_full.csv | Measure-Object | Select-Object -expand count
$absent = Import-Csv .\Absent_Output_$date.csv | Measure-Object | Select-Object -expand count
$present = $total-$absent

$index = 0

$list = Import-Csv .\Absent_Output_$date.csv -DeLimiter "," |
Select-Object 'Name','Year' | 
ConvertTo-Csv -NoTypeInformation | ForEach-Object { "{0}){1}" -f $index++, $_ } | ForEach-Object { $_ -replace '"', ""} | ForEach-Object { $_ -replace ',', " "} | ForEach-Object { $_ -replace 'WSC', "(WSC"} | ForEach-Object { $_ -replace '08', "08)"} | Out-File temp.txt


$output = "WSC 108, 208 & 308 - Karate 1,2,3 ($date2) 
$present/$total
*Tidak Hadir/Absence:* "
$output | Add-Content -Path .\Absent_output_$date.txt

#Generate output file
echo ""
echo "Generating Absentee Output File..."
Add-Content -Path .\Absent_output_$date.txt -Value (Get-Content ".\temp.txt" | Select-Object -Skip 1)

#Remove temp items
echo ""
echo "Removing temporary files..."
Remove-Item .\temp.txt
Remove-Item .\Absent_Output_$date.csv
Remove-Item .\"Attendance Form USM Karate ($date).csv.zip"
Remove-Item .\"Attendance Form USM Karate ($date).csv"

notepad.exe .\Absent_Output_$date.txt

echo "Run completed"