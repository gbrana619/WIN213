<#
.SYNOPSIS
   Manage client's information
.DESCRIPTION
   This module is used to manage client's information of Canadian Tire, 
   based on the database - clientinfo.cvs. 
   There are 7 options for selection.
.NOTES
   Group Member Names: Zijun Liu; Joy Sikder; Gyan Bahadur Rana
   DateLastModified: 07-28-2020
 #>

########################################################################################
#This function is used to remove duplicate data from database file clientinfo.csv
########################################################################################

# Create Remove-Duplicates Function.
Function Remove-Duplicates {

# Set the file location in clientinfo folder which is same name as script file name.
$FilePath = "$Home\Documents\WindowsPowerShell\Modules\clientinfo\clientinfo.csv"

# Test if the path to the database file exists.
$Result = Test-Path $FilePath

If ($Result) {

 $CSV = Import-Csv $FilePath

# Test to see if Duplicate first and last names exist in the database.
 $Duplicate = $CSV | Select-Object FirstName,LastName | Group-Object FirstName, LastName | ? {$_.Count -gt 1}

 If ($Duplicate -gt 0) {
 $Count = $Duplicate.Count }
 Else {
   $Count = 0}
  

# Display Title for Duplicate Data.
 Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White
 Write-host "         There are" -NoNewline
 Write-host " $Count Duplicates" -NoNewline -ForegroundColor Red
 Write-host " in Databse file"
 Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White
 Write-Output "`n"
 Write-host "Remove Duplicates?.Please CTRL + C to exit or " -NoNewline
 Pause

# Remove the Duplicate Data and save to a new csv file.
 $CSVNew = $FilePath + ".new"
 $CSV | Sort-Object * -Unique | Export-Csv $CSVNew -NoTypeInformation

# Delete the original csv file.
 Remove-Item $FilePath

# Rename the new csvfile back to the original one.
 Move-Item $CSVNew $FilePath
 If ($Count -gt 0) {
    "`n"
    Write-host "Removed Duplicates" -ForegroundColor DarkGreen -BackgroundColor White
 }

} Else {
    Write-host "WARNING: File" -Foregroundcolor Yellow
    Write-host " $FilePath does not exist. Check your path and try again." -Foregroundcolor Yellow
    Pause
    Return
  }
} #end of Remove-Duplicates

########################################################################################
#This function is used to check if input client's name exists or not 
########################################################################################

# Create Get-ClientName Function.
Function Get-ClientName {

$FilePath = "$Home\Documents\WindowsPowerShell\Modules\clientinfo\clientinfo.csv"

# Test if the path to the database file exists.
$Result = Test-Path $FilePath

If ($Result) {
    $ClientName = Read-Host "Enter client name [Jane Doe]?"
    $NewClientName = $ClientName.Split(" ")
    $Global:FirstName = $NewClientName[0]
    $Global:LastName = $NewClientName[1]

    $CSV = Import-Csv $FilePath
    $CSVName = $CSV | Where-Object {$_.FirstName -eq $FirstName -and $_.LastName -eq $LastName }
    
    If ( $CSVName -ne $Null){
    Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White
               "             Client is in Databas file"
    Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White

    } Else {
    Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White
    Write-host "             Client is " -NoNewline
    Write-host " not" -NoNewline -ForegroundColor Red
    Write-host " in Databas file"
    Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White
      }
}
Else {
    Write-host "WARNING: File" -Foregroundcolor Yellow
    Write-host " $FilePath does not exist. Check your path and try again." -foregroundcolor Yellow
    Pause
  }

} #end of Get-ClientName

########################################################################################
#This function is used to update client information
########################################################################################

# Create Set-ClientInfo Function.
Function Set-ClientInfo {

$FilePath = "$Home\Documents\WindowsPowerShell\Modules\clientinfo\clientinfo.csv"

# Test if the path to the database file exists.
$Result = Test-Path $FilePath
If ($Result) {
   Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White
               " Editing Client $Global:FirstName $Global:LastName Information"
   Write-host "=============================================================" -ForegroundColor DarkGreen -BackgroundColor White
   "`n"

   $CSV = Import-Csv $FilePath
   $CSVName = $CSV.FirstName
   ##[array]::IndexOf($CSVName,"$FirstName")

   # Update client information.
   $NFirstName = Read-Host "Enter first name "
   $NLastName = Read-Host "Enter last name "
   $NBusiness = Read-Host "Enter business type "
   $NTelephone = Read-Host "Enter telephone number "
   $NEmail = Read-Host "Enter email address "

   $CSVNew = $FilePath + ".New"
   $CSV | ForEach-Object {
   If ($_.FirstName -eq $FirstName) {
   
       $_.FirstName = $NFirstName
       $_.LastName = $NLastName
       $_.Business = $NBusiness
       $_.Telephone = $NTelephone
       $_.Email = $NEmail
   }
   $_ | Export-Csv $CSVNew -NoTypeInformation -Append }

   # Delete the original csv file.
   Remove-Item $filepath

   # Rename the new csvfile back to the original one.
   Move-Item $CSVNew $FilePath

   "`n"
   Write-host "Client information edited" -ForegroundColor DarkGreen -BackgroundColor White
}


} #end of Set-ClientInfo

########################################################################################
#This is the menu funciton and uses to display menu options
########################################################################################

# Create Get-Menu Function.
Function Get-Menu {
<#....#>
$Menu = @"

╔═════════════════════════════════════════╗
║  Database Management -- Win213Assign04  ║
╠═════════════════════════════════════════╣
║                                         ║
║        1. Remove Duplicates             ║
║        2. Search by Client Name         ║
║        3. Edit Client Information       ║
║        4. Delete Client Information     ║
║        5. Add Client Information        ║
║        6. Show Client Information       ║
║        7. Exit Program                  ║
║                                         ║
╚═════════════════════════════════════════╝
"@

# Menu Interaction using Read-Host and Switch.
Do {

# Clear the screen and display menu.
Clear-Host
Write-Host "$Menu" -ForegroundColor Green -BackgroundColor Black

# User can input number 1-7 for option selection.
$Selection = Read-Host "Enter a selection [1-7]"
Write-Output "`n`n"

Switch ($Selection) 
    {
  
      '1' {Remove-Duplicates;Pause;Break}
      '2' {Get-ClientName;Pause;Break}
      '3' {Set-ClientInfo;Pause;Break}
      '4' {Remove-ClientInfo;Pause;Break}
      '5' {Add-ClientInfo;Pause;Break}
      '6' {Show-ClientInfo;Pause;Break}
      '7' {Exit;Break}
      Default {"Please make a valid choice.";Pause;Break}
    }
} Until ($Selection -eq '7')
} #end of Get-Menu

# Create Show-Menu Function.
Function Show-Menu {
  Get-Menu
} #end of Show-Menu

########################################################################################
