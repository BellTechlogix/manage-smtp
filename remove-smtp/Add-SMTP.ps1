<#
Remove-SMTP.ps1
Created By - Kristopher Roy
Ceated On - 11/30/2017
Last Modified - 12/5/2017
Purpose - The purpose of this script is to loop through a list of users and remove an SMTP address from their list, initially based upon csv list, but eventually will include options to manually input users, or grab and report on all!
#>


#import the AD Module
Import-Module ActiveDirectory

#Function to select a file
function Get-FileName
{
  param(
      [Parameter(Mandatory=$false)]
      [string] $Filter,
      [Parameter(Mandatory=$false)]
      [switch]$Obj,
      [Parameter(Mandatory=$False)]
      [string]$Title = "Select A File"
    )
 
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.FileName = $Title
  #can be set to filter file types
  IF($Filter -ne $null){
  $FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
	$OpenFileDialog.filter = $FilterString}
  if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
  $OpenFileDialog.filter = $Filter
  }
  $OpenFileDialog.ShowDialog() | Out-Null
  ## dont bother asking, just give back the object
  IF($OBJ){
  $fileobject = GI -Path $OpenFileDialog.FileName.tostring()
  Return $fileObject
  }
  else{Return $OpenFileDialog.FileName}
}

#Select and then import your csv list
$file = get-filename -Filter "csv" -Title "Select User List" -obj
$userlist = import-csv $file.FullName

#SMTP Alias domain to add
$addalias = "@studentconnections.org"

$progresscount = 0
FOREACH($User in $Userlist)
{
	$progresscount++
	#Display Progress Bar
	Write-Progress -Activity ("Adding $addalias SMTP Addresses . . ."+$user.sAMAccountName) -Status "Scanned: $progresscount of $($userlist.sAMAccountName.Count)" -PercentComplete ($progresscount/$userlist.sAMAccountName.Count*100)
	
	#convert email to username
	$username = $user.emailaddress.split("@")[0]
	

	$aduser = Get-ADUser $user.sAMAccountName -Properties proxyaddresses
    $aduser.ProxyAddresses.add("SMTP:$($username+$addalias)")
    $result = Set-ADUser -Instance $aduser
    $username = $null
    $aduser = $null
	$result = $null
}