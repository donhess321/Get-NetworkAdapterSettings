<#
    .SYNOPSIS
    Get all network settings for a computer.

    .DESCRIPTION
    Get all network settings for a computer with multiple options to
    return objects or export to a file.

    .RETURNS
    Network adapter configuration objects from the remote machine or
    A DataTable of all configuration information

    .OUTPUTS
    An HTML of the DataTable (the default) (which you can manually change the extension to .xls and open in MS Excel)
    A CSV file
    A text file in Format-List formating

    .PARAMETER Computers
    Array:  List of computer names to log into via PS Remoting

    .PARAMETER OutputOriginalObjects
    Boolean: Output the network adapter configuration objects to the pipeline

    .PARAMETER OutputDataTable
    Boolean: Output a DataTable of the network adapter configuration objects to the pipeline
    
    .PARAMETER ExportHtml
    Boolean: Export all data to an HTML file

    .PARAMETER ExportCsv
    Boolean: Export all data to an CSV file

    .PARAMETER ExportList
    Boolean: Export all data to an list formated text file

    .PARAMETER OutputFileName
    String:  File name for output.  This can be short or full name format.
    Do NOT add an extension as the correct one will be added before export.

    .NOTES
    This is intended to be run over PS Remoting.
    Booleans were used instead of switches so the true/false can be set for defaults
    Multiple output file types can be specified simultaneously

    Author: Donald Hess
    Version History:
        2.1  2018-05-30  Added line to merge $aComputers if something is passed in 
        2.0  2018-04-23  Added multiple output functionality, add GUI name to adapter
        1.0  2017-07     Release
    
    .EXAMPLE
    Get-NetworkAdapterSettings -Computers 'comp1','comp2','etc' -ExportHtml $true -OutputFileName ''
    .EXAMPLE
#> 

param ( [array] $Computers = @(), 
        [bool] $OutputOriginalObjects = $false,
        [bool] $OutputDataTable = $false,
        [bool] $ExportHtml = $true,
        [bool] $ExportCsv = $false,
        [bool] $ExportList = $false,
        [string] $OutputFileName = 'netadptr_results'  # No Extension, can be short or full path
      )
Set-StrictMode -Version latest -Verbose
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction']='Stop'

if ( -not $Computers ) {
    # This will filter so we get only computer object that are backed by a real workstation
    # Update to your environment to get the workstations you want
    $sDomainSnip = (Get-WmiObject Win32_ComputerSystem).Domain.Trim().Substring(0,4)    
    $aComputers = Get-ADComputer -Filter * | Select-Object Name,DNSHostName,Enabled | `
        Where-Object { $_.Enabled -and $_.Name -Like "$sDomainSnip*" -and $null -ne $_.DNSHostName } | `
        ForEach-Object { $_.Name } | Sort
} else {
    # Something passed in
    $aComputers = $Computers
}
$sb1 = {
Write-Host 'Working on:' $env:COMPUTERNAME
# Need to get the Name you see in the GUI to be able to use with tools like netsh
# https://github.com/devops-collective-inc/powershell-networking-guide/blob/master/manuscript/identifying-network-adapters.md
$htNetAdaptersSettings = @{}
Get-WmiObject Win32_NetworkAdapterSetting | Foreach-object {
    [wmi] $_.element | Select NetConnectionId, Name, DeviceId, InterfaceIndex, NetconnectionStatus | `
    ForEach-Object { $htNetAdaptersSettings.Add([int] $_.DeviceId,$_) }
}
Get-WmiObject Win32_NetworkAdapterConfiguration | `
Where-Object { ($_.Description -NotLike "*Miniport*") } | `
Where-Object { ($_.Description -NotLike "*ISATAP*") } | `
Where-Object { ($_.Description -NotLike "*Debug*") } | `
Select Description,
    @{Name="Index"; Expression={[int] $_.Index}}, 
    @{Name="Name"; Expression={$htNetAdaptersSettings[[int] $_.Index].NetConnectionId}}, 
    @{Name="DHCPEnabled"; Expression={$_.DHCPEnabled}}, 
    @{Name="DHCPLeaseObtained"; Expression={$_.DHCPLeaseObtained}},
    @{Name="DHCPServer"; Expression={$_.DHCPServer}},
    @{Name="DNSDomain"; Expression={$_.DNSDomain}},
    @{Name="DNSDomainSuffixSearchOrder"; Expression={$_.DNSDomainSuffixSearchOrder}},
    @{Name="DNSEnabledForWINSResolution"; Expression={$_.DNSEnabledForWINSResolution}},
    @{Name="DNSHostName"; Expression={$_.DNSHostName}},
    @{Name="DNSServerSearchOrder"; Expression={$_.DNSServerSearchOrder}},
    @{Name="DomainDNSRegistrationEnabled"; Expression={$_.DomainDNSRegistrationEnabled}},
    @{Name="FullDNSRegistrationEnabled"; Expression={$_.FullDNSRegistrationEnabled}},
    @{Name="IPAddress"; Expression={$_.IPAddress}},
    @{Name="IPSubnet"; Expression={$_.IPSubnet}},
    @{Name="DefaultIPGateway"; Expression={$_.DefaultIPGateway}},
    @{Name="WINSEnableLMHostsLookup"; Expression={$_.WINSEnableLMHostsLookup}},
    @{Name="WINSHostLookupFile"; Expression={$_.WINSHostLookupFile}},
    @{Name="WINSPrimaryServer"; Expression={$_.WINSPrimaryServer}},
    @{Name="WINSScopeID"; Expression={$_.WINSScopeID}},
    @{Name="WINSSecondaryServer"; Expression={$_.WINSSecondaryServer}},
    @{Name="MACAddress"; Expression={$_.MACAddress}},
    @{Name="MTU"; Expression={$_.MTU}},
    @{Name="TcpipNetbiosOptions"; Expression={$_.TcpipNetbiosOptions}},
    @{Name="TcpWindowSize"; Expression={$_.TcpWindowSize}},
    @{Name="ServiceName"; Expression={$_.ServiceName}}
}
function funcConvertTo-DataTable {
    <#  .SYNOPSIS
            Convert regular PowerShell objects to a DataTable object.
        .DESCRIPTION
            Convert regular PowerShell objects to a DataTable object.
        .EXAMPLE
            $myDataTable = $myObject | ConvertTo-DataTable
        .NOTES
            Name: ConvertTo-DataTable
            Author: Oyvind Kallstad @okallstad
            Version: 1.1
    #>
    [CmdletBinding()]
    param (
        # The object to convert to a DataTable
        [Parameter(ValueFromPipeline = $true)]
        [PSObject[]] $InputObject,

        # Override the default type.
        [Parameter()]
        [string] $DefaultType = 'System.String'
    )
    begin {
        # Create an empty datatable
        try {
            $dataTable = New-Object -TypeName 'System.Data.DataTable'
            Write-Verbose -Message 'Empty DataTable created'
        } catch {
            Write-Warning -Message $_.Exception.Message
            break
        }
        # Define a boolean to keep track of the first datarow
        $first = $true
        # Define array of supported .NET types
        $types = @(
            'System.String',
            'System.Boolean',
            'System.Byte[]',
            'System.Byte',
            'System.Char',
            'System.DateTime',
            'System.Decimal',
            'System.Double',
            'System.Guid',
            'System.Int16',
            'System.Int32',
            'System.Int64',
            'System.Single',
            'System.UInt16',
            'System.UInt32',
            'System.UInt64'
        )
    }
    process {
        # Iterate through each input object
        foreach ($object in $InputObject) {
            try {
                # Create a new datarow
                $dataRow = $dataTable.NewRow()
                Write-Verbose -Message 'New DataRow created'
                # Iterate through each object property
                foreach ($property in $object.PSObject.get_properties()) {
                    # Check if we are dealing with the first row or not
                    if ($first) {
                        # handle data types
                        if ($types -contains $property.TypeNameOfValue) {
                            $dataType = $property.TypeNameOfValue
                            Write-Verbose -Message "$($property.Name): Supported datatype <$($dataType)>"
                        } else {
                            $dataType = $DefaultType
                            Write-Verbose -Message "$($property.Name): Unsupported datatype ($($property.TypeNameOfValue)), using default <$($DefaultType)>"
                        }
                        # Create a new datacolumn
                        $dataColumn = New-Object 'System.Data.DataColumn' $property.Name, $dataType
                        Write-Verbose -Message 'Created new DataColumn'

                        # Add column to DataTable
                        $dataTable.Columns.Add($dataColumn)
                        Write-Verbose -Message 'DataColumn added to DataTable'
                    }                  
                    # Add values to column
                    if ($property.Value -ne $null) {
                        # If array or collection, add as XML
                        if (($property.Value.GetType().IsArray) -or ($property.TypeNameOfValue -like '*collection*')) {
                            $dataRow.Item($property.Name) = $property.Value | ConvertTo-Xml -As 'String' -NoTypeInformation -Depth 1
                            Write-Verbose -Message 'Value added to row as XML'
                        } else {
                            $dataRow.Item($property.Name) = $property.Value -as $dataType
                            Write-Verbose -Message "Value ($($property.Value)) added to row as $($dataType)"
                        }
                    }
                }
                # Add DataRow to DataTable
                $dataTable.Rows.Add($dataRow)
                Write-Verbose -Message 'DataRow added to DataTable'
                $first = $false
            } catch {
                Write-Warning -Message $_.Exception.Message
            }
        }
    }
    end { Write-Output (,($dataTable)) }
} # End funcConvertTo-DataTable

$aResults = @(Invoke-Command -ThrottleLimit 1 -ScriptBlock $sb1 -ComputerName $aComputers -ErrorAction Continue)

if ( $OutputOriginalObjects ) {
    $aResults  # Returning each object, not array
}
if ( $OutputDataTable -or $ExportHtml ) {
    $aDtResults = @()
    $aResults | ForEach-Object {
        $aDtResults += @(funcConvertTo-DataTable $_)
    }
}
if ( $OutputDataTable ) {
    $aDtResults
}
if ( $ExportHtml ) {
    #HTML output
    $aHtmlContent = @()
    $aHtmlContent += '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml"><head><title>HTML TABLE</title>
    <style type="text/css">
    table {
	    border: thin solid lightgray;
	    border-collapse: collapse;
    }
    td {
	    border: thin solid lightgray;
	    padding-left: 10px;
	    padding-right: 10px;
    }
    </style></head><body>'
    $aDtResults | ForEach-Object {
        $aHtmlContent += ($_ | Select * -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-Html -Fragment) -join ''
    }
    $aHtmlContent += '</body></html>'
    $aHtmlContent -join '</br><hr></br>' > (@($OutputFileName,'.html') -join '')
}
if ( $ExportCsv ) {
    $aResults | Export-Csv -NoTypeInformation -Path (@($OutputFileName,'.csv') -join '')
}
if ( $ExportList ) {
    $aResults | Out-File -Force -Append -FilePath (@($OutputFileName,'.txt') -join '')

}
