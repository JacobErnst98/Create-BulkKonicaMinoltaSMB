. ..\Dependencies\Choose-Ou.ps1 # https://raw.githubusercontent.com/ITMicaH/Powershell-functions/master/Active-Directory/OUs/ChooseADOrganizationalUnit.ps1

$hostname = "<Server>" # Hostname for File Server 
$saName = "<Username>" # User With permission to scan folder
$scannerIP = "10.100.0.200" # IP of KM MFP
$folderPath1 = "Users\"
$folderPath2 = "\Data\Documents\Scans"

write-host "Password for $saName ?"
$saPassword = read-host
$ou = Choose-ADOrganizationalUnit
$users = Get-ADUser -Filter * -SearchBase $ou.DistinguishedName -Properties samAccountName
$users = $users | Out-GridView -Title "Hold CTRL to select users to add" -PassThru

& "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://$($scannerIP):30092/goform/admin_en.html"
foreach ($user in $users){
$Fullname = $user.Name
$shortName =$user.Name.Split(" ")[0]  + " "+ $user.Name.Split(" ")[1][0]+"."
$filePath= $folderPath1 + $user.samaccountname + $folderPath2


$htmlstring = @"
<form method="POST" action="https://$scannerIP:30092/goform/ipscan_en.html">
<table width="95%" border="10" align="CENTER">
<tbody><tr>
<td class="title">Register Name</td>
<td class="celldef"><input type="text" name="entry----" size="50" maxlength="24" value="$Fullname">(Maximum: 24 Characters)</td>
</tr>
<tr>
<td class="title">Reference Name</td>
<td class="celldef"><input type="text" name="search---" size="50" maxlength="24" value="$shortName" >(Maximum: 24 Characters)</td>
</tr>
<tr>
<td class="title">Host Address</td>
<td class="celldef"><input type="text" name="smbserv__" size="50" maxlength="253" value="$hostname" >(Maximum: 253 Characters)</td>
</tr>
<tr>
<td class="title">File Path</td>
<td class="celldef"><input type="text" name="smbdirn__" size="50" maxlength="255" value="$filePath" >(Maximum: 255 Characters)</td>
</tr>
<tr>
<td class="title">Login Name</td>
<td class="celldef"><input type="text" name="smbuser__" size="50" maxlength="32" value="$saName" >(Maximum: 32 Characters)</td>
</tr>
<tr>
<td class="title">Password</td>
<td class="celldef"><input type="password" name="smbpass__" size="50" maxlength="32" value="$saPassword" >(Maximum: 32 Characters)</td>
</tr>
<tr>
<td class="title">Daily Use Registration</td>
<td class="celldef"><input type="checkbox" name="common___" value="on_______" checked >(Registered as daily use when checked)</td>
</tr>
</tbody></table>
<input type="hidden" name="type_____" value="smbpage__">
<input type="hidden" name="command__" value="addcmd___">
<input type="submit" value="Registration">
</form>

"@

$htmlstring | Out-File -FilePath $($Fullname + ".html" )

$cDir = $PWD.Path.Split("::")[2].replace("\","/").replace("//","")

& "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "File://$cDir/$($Fullname + ".html" )"
}
Write-Host "Press Enter When your all done."
Read-Host | null
