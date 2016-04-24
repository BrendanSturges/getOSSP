<#

Script created by Brendan Sturges, reach out if you have any issues.
This script queries a file the user chooses and checks all servers within for the OS & Service Pack

#>

Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

Function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension

}

#Open a file dialog window to get the source file
$serverList = Get-Content -Path (Get-FileName)

#open a file dialog window to save the output
$fileName = Save-File $fileName
$errorActionPreference = "SilentlyContinue"
$i = 0
foreach($server in $serverList) {
	Try {
		$domain = ''
		$reader = ''
		$OS = ''
		$SP = ''
		$ErrorMessage = ''
	
		$pingAndReturn = (ping $server -n 1).split('.')
		$domain = $pingAndReturn[2].toUpper()
		$reader = Get-WMIObject -computername $server -class win32_operatingsystem | Select-Object -Property @{n='HostName';e={$_.CSName}},@{n='OSDescription';e={$_.Caption}},@{n='ServicePack';e={$_.ServicePackMajorVersion}}
		
		$OS = $reader | select -expandproperty OSDescription
		$SP = $reader | select -expandproperty ServicePack
		
		$props = [ordered]@{
			'Server' = $server
			'Operating System' = $OS.toString()
			'Service Pack' = $SP.toString()
			'Domain' = $domain
			'Details' = ''
			
		} 
		$obj = New-Object -TypeName PSObject -Property $props
	}

	Catch {
		$ping = Test-Connection -ComputerName $server -Count 2 -Quiet
		if($ping)
			{
			$ErrorMessage = $_.Exception.Message
			}
		else
			{
			$ErrorMessage = 'Server is Offline'
			}
		
		$props = [ordered]@{
			'Server' = $server
			'Operating System' = ''
			'Service Pack' = ''
			'Domain' = $domain
			'Details' = $ErrorMessage
			}
			
		$obj = New-Object -TypeName PSObject -Property $props
	
	}
	Finally {
		$data = @()
		$data += $obj
		$data | Export-Csv $fileName -noTypeInformation -append	
	}
	$i++
	Write-Progress -activity "Checking server $i of $($serverList.count)" -percentComplete ($i / $serverList.Count*100)	
}


