<#
This Program goes through all the remote pcs iteratively and adds files, runs programs, or deletes things
You need to make sure that you are running powershell as an ADMIN or else it will not work
If the password to enter the virtual machines or remote pcs is changed you will need to change it in the two password sections 
Make sure powershell is enabled. The command for that is: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypas
Create an installer package (.mst or .msi) and place that in the $installCommand and change $softwareName
Bypass dual verification regedit -> HKEY_LOCAL_MACHINE -> Software -> Microsoft -> Terminal Server Client -> New -> DWORD
Value name: AuthenticationLevelOverride Value Data: 0
#>

# Define the path to the shared file if you want to copy or delete something on multiple PCs
$sharedFilePath = "\\remotepcaddress\e$\Software\installer_package.msi"

# This loops through 192.100.9.1 to 192.100.9.99
# Base IP Address
$baseIP = "192.100.9."
# Starting IP address 
$counter = 1 
# Ending IP address
$endCounter = 99

# Put the path to the installer package here
$registryPath = "HKLM:\\HKEY_LOCAL_MACHINE\Software\Windows\CurrentVersion\"
$softwareName = "Software"

# Iterate through each PC 
while ($counter -le $endCounter) {
	# Construct the IP address using concatenation 
	$IPAddress = $baseIP + $counter
	# Construct the copypath
	$CopyDeletePath = "\\$IPAddress\c$\Users\Desktop"
	$runpath = "\\$IPAddress\c$\Users\Desktop\installer_package.msi"

    Write-Host "Connecting to $IPAddress"

try{
	# Start Remote Session (change pass:PASSWORD to the remote PC password)
	cmdkey /generic:TERMSERV/$IPAddress /user:$env:USERNAME /pass:abc123456
	mstsc /v:$IPAddress
        
	# Types in password
	Add-Type -AssemblyName System.Windows.Forms
	Start-Sleep -Seconds 15
	# Replace the first SendWait with the password
	[System.Windows.Forms.sendKeys]::SendWait("abc123456")
	[System.Windows.Forms.sendKeys]::SendWait("{ENTER}")
  Start-Sleep -Seconds 5
    

	$rdcWindow = Get-Process | Where-Object {$_.MainWindowTitle -eq "$IPAddress" - Remote Desktop Connection}

	if($rdcWindow -eq $null){
		Write-Host "Couldn't connect to $IPAddress"
		goto Exit
	}else{
		# Copies Item to Remote PC
       	Copy-Item -Path $sharedFilePath -Destination $CopyDeletePath -Recurse
		# Delete item from the remote PC
        # Get-ChildItem -Path $CopyDeletePath -Remove-Item -Force -Confirm:$False

		# Run the first program on the remote PC
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $runPath /quiet" -Wait 

		# Wait for installation to finish
		while($true){
			if(Get-Children $registryPath | Where-Object {$_.GetValue("DisplayName") -eq $softwareName}){
				Write-Host "$softwareName installed on $IPAddress successfully"
				break
			} else{
				Start-Sleep -Seconds 15
			}
		}
	}
	
} catch {
	Write-Host "Error on: $IPaddress "
} finally {
	:Exit
    # Find Process Associated with the current Remote PC through task manager
	$rdcProcess = Get-Process -Name "mstsc"
	# Close Window
	$rdcProcess | Stop-Process
	# Increment counter 
    $counter++
	Start-Sleep -Seconds 5
}
}
