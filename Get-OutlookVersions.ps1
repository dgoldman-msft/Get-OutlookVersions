Function Get-OutlookVersions {
	
    <#
		.SYSNOPSIS
            Find outlook build numbers
	
        .DESCRIPTION
            Search domain computers based on filter and find all outlook build numbers
	
        .PARAMETER Filter
            Search filter to narrow down search results
	
        .PARAMETER NeedsAdminCreds
            Switch used to indicate we need to connect using credentials

        .PARAMETER NeedsADCreds
            Switch used to indicate we need to connect using credentials to Active Directory

        .PARAMETER Batch
            Connect to a single or subset of machines without using Get-ADComputer

        .PARAMETER ShowResults
            Show the results to the console without exporting

        .PARAMETER ShowFailures
            Display the failures to the console

        .PARAMETER ExportResults
            Specify to export the results to disk

        .PARAMETER ExportFailures
            Specify to export the failures to disk
        
        .PARAMETER ExportPath
            Path to exported location

        .PARAMETER ExportResultsPath
            Path to save the exported results file

        .PARAMETER ExportFailuresPath
            Path to save the exported failures file

        .PARAMETER AuthType
            Specifies the authentication method to use. The acceptable values for this parameter are: Negotiate or 0 Basic or 1. The default authentication method is Negotiate.

		.EXAMPLE
			Get-OutlookVersions
			
			Run with no credentials and find all clients with outlook and get the build number

		.EXAMPLE
			Get-OutlookVersions -Verbose

			Run with no credentials in verbose mode and find all clients with outlook and get the build number

		.EXAMPLE
			Get-OutlookVersions -NeedsAdminCreds

			Run with credentials and find all clients with outlook and get the build number

        .EXAMPLE
			Get-OutlookVersions -Batch "Machine01"

			Run the function against a single machine

        .EXAMPLE
			Get-OutlookVersions -Batch "Machine01", "Machine02" -Export

			Run the function against a subset of machines and export to c:\OutlookBuilds.csv

        .NOTES
            For more information on the filters that can be used please reference: https://docs.microsoft.com/en-us/powershell/module/activedirectory/get-adcomputer?view=windowsserver2019-ps
    #>
	
    [cmdletBinding()]
    Param(
        [string]
        $Filter = '*',
		
        [switch]
        $NeedsAdminCreds,

        [switch]
        $NeedsADCreds,

        [switch]
        $ExportResults,

        [switch]
        $ExportFailures,

        [switch]
        $ShowResults,

        [switch]
        $ShowFailures,

        [object]
        $Batch,

        [Int32]
        $AuthType = 0,

        [string]
        $ExportPath = [environment]::GetFolderPath("Desktop"),

        [string]
        $ExportResultsPath = (Join-Path -Path $ExportPath -ChildPath "OutlookBuildsResults.csv"),

        [string]
        $ExportFailuresPath = (Join-Path -Path $ExportPath -ChildPath "OutlookBuildsFailures.csv")
    )
	
    begin {
        Write-Host "Starting…"
        $StartTime = $(get-date)

        If ($NeedsAdminCreds.IsPresent) { 
            Write-verbose "Obtaining credentials"
            $ADAdminCreds = Get-Credential 
        }

        Update-TypeData -TypeName OutlookInfo -DefaultDisplayPropertySet Product, Version, ComputerName -DefaultDisplayProperty DisplayName -DefaultKeyPropertySet CustomProperties -Force
        Update-TypeData -TypeName FailureInfo -DefaultDisplayPropertySet Reason, Computer, Exception -DefaultDisplayProperty Message -DefaultKeyPropertySet CustomProperties -Force
    }
	
    process {
        $exceptionRecords = @()
        $results = @()
        $computersFound = 0
        $exceptionsFound = 0
		
        Try {
            if ($NeedsADCreds.IsPresent) {
                # Filter can be used to search for any search pattern we need to filter down the results
                $computers = Get-ADComputer -Filter $Filter -AuthType $AuthType -Credential $ADAdminCreds | Select-Object Name
            }
            elseif ($batch) { $computers = $batch }
            else {
                # Filter can be used to search for any search pattern we need to filter down the results
                $computers = Get-ADComputer -Filter $Filter -AuthType $AuthType | Select-Object Name
            }
        }
        Catch {
            $_.Exception.Message
            return
        }
		
        foreach ($computer in $computers) {
            $computersFound ++
            
            if (-NOT ($Batch)) { $computer = $computer.Name }

            Write-verbose "Scanning $($computer) for Outlook builds"
			
            if ($NeedsAdminCreds) {
                $found = Invoke-Command -ComputerName $Computer -ScriptBlock { 
                    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail* } -ErrorAction SilentlyContinue -ErrorVariable Failed -Credentials $creds
            }
            else {
                $found = Invoke-Command -ComputerName $Computer -ScriptBlock {
                    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail* } -ErrorAction SilentlyContinue -ErrorVariable Failed
            }

            if ($found) { 
                $machineData = [PSCustomObject]@{
                    PSTypeName   = 'OutlookInfo'
                    Product      = $found.DisplayName
                    Version      = $found.DisplayVersion
                    ComputerName = $found.PSComputerName
                }
                
                $results += $machineData
            }

            # Save the error record
            if ($Failed) {
                $exceptionsFound ++
                Write-verbose "Saving error record for review"
                
                $exception = [PSCustomObject]@{
                    PSTypeName = 'FailureInfo'
                    Reason     = $Failed.FullyQualifiedErrorId
                    Computer   = $Failed.TargetObject
                    Exception  = [string]$Failed.Exception.Message
                }
                
                $exceptionRecords += $exception
            }
        }

        if ($ExportResults -and $computersFound -gt 0) {
            Write-Verbose "Exporting data to $($ExportPath)"
            try {
                $results | Export-Csv -Path $ExportResultsPath -NoTypeInformation -ErrorAction Stop
            }
            catch {
                $_.Exception.Message
                return
            }
            Write-Verbose "Export complete!"
        }

        if ($ExportFailures -and $exceptionsFound -gt 0) {
            Write-Verbose "Exporting data to $($ExportFailuresPath)"
            try {
                [PSCustomObject]$exceptionRecords | Export-Csv -Path $ExportFailuresPath -NoTypeInformation -ErrorAction Stop 
            }
            catch {
                $_.Exception.Message
                return
            }
            Write-Verbose "Export complete!"
        }

        $elapsedTime = $(get-date) - $StartTime
        Write-Host "Completed…`r"
        "`rTotal run time: {0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)

        # User specified - report the results found
        if ($ShowResults) { $results }
        
        # User specified - report the errors found
        if ($ShowFailures) { $exceptionRecords | Format-Table -AutoSize -Wrap}
        else { Write-Verbose "No failures detected!" }
            
        Write-Host "Total computers scanned: $($computersFound)"
        Write-Host "Total connection failures: $($exceptionsFound)`r`n"
    }
	
    end {}
}
