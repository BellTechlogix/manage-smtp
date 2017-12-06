<#
Remove-SMTP.ps1
Created By - Kristopher Roy
Ceated On - 11/29/2017
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

#SMTP Alias domain to remove
$removealias = "@studentconnections.org"

$progresscount = 0
FOREACH($User in $Userlist)
{
	$progresscount++
	#Display Progress Bar
	Write-Progress -Activity ("Removing $removealias SMTP Addresses . . ."+$user.EmailAddress) -Status "Scanned: $progresscount of $($userlist.emailaddress.Count)" -PercentComplete ($progresscount/$userlist.emailaddress.Count*100)
	
	#convert email to username
	$username = $user.emailaddress.split("@")[0]
	
	#Get ADUser object
	$aduser = Get-ADUser $username -Properties proxyaddresses

    #Check if SMTP is Primary or alt
	ForEach ($proxy in $aduser.ProxyAddresses) {
        If ($proxy.StartsWith("SMTP:")) {
            If ($proxy -eq "SMTP:$($username+$removealias)") {
                $blnPrimaryOld = $true
            }
        }
        ElseIf ($proxy.StartsWith("smtp:")) {
            If ($proxy -eq "smtp:$($username+$removealias)") {                
                $blnAliasOld = $true
            }
        }
    }

	#Run if SMTP is Primary
	IF($blnPrimaryOld -eq $true)
	{
		$aduser.ProxyAddresses.remove("SMTP:$($username+$removealias)")
		$result = Set-ADUser -Instance $aduser
	}

	#Run if SMTP is alt
	IF($blnAliasOld -eq $true)
	{
		$aduser.ProxyAddresses.remove("smtp:$($username+$removealias)")
		$result = Set-ADUser -Instance $aduser
	}

	#clear vars before next loop
	$username = $null
	$aduser = $null
	$result = $null
}