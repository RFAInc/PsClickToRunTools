function Get-WebRequestTable {
    <#
    .SYNOPSIS
    Attempts to scrape table from webpage.
    .DESCRIPTION
    Scrapes a given numbered table for the provided Web 
    Request response from the Invoke-WebRequest cmdlet.
    .PARAMETER WebRequest
    HtmlWebResponseObject returned from Invoke-WebRequest cmdlet. 
    .PARAMETER TableNumber
    Index number of the table on the page, in order. First table is default.
    .EXAMPLE
    $r = Invoke-WebRequest $url
    Get-WebRequestTable $r -TableNumber 0 | Format-Table -Auto

    P1              P2         P3                   P4
    --              --         --                   --
    Gardiner Number Hieroglyph Description of Glyph Details
    Q1                         Seat                 Phono. st, ws, . In st ?seat, place,? wsir ?Osiris,? ?tm ?perish.?
    Q2                         Portable seat        Phono. ws. In wsir ?Osiris.?
    Q3                         Stool                Phono. p.
    Q4                         Headrest             Det. in wrs ?headrest.?
    Q5                         Chest                Det. in hn ?box,? ?fdt ?chest.?
    Q6                         Coffin               Det. or Ideo. in qrs ?bury,? krsw ?coffin.?
    Q7                         Brazier with flame   Det. of fire. In ?t ?fire,? s?t ?flame,? srf ?temperature.?
    .NOTES
    From https://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/
    #>
    param(
        [Parameter(Position=0,Mandatory = $true)]
        [Microsoft.PowerShell.Commands.HtmlWebResponseObject]
        $WebRequest,

        [Parameter()]
        [int]
        $TableNumber = 0
    )

    ## Extract the tables out of the web request
    $tables = @($WebRequest.ParsedHtml.getElementsByTagName("TABLE"))
    $table = $tables[$TableNumber]
    $titles = @()
    $rows = @($table.Rows)

    ## Go through all of the rows in the table
    foreach($row in $rows) {
        $cells = @($row.Cells)
        ## If we've found a table header, remember its titles
        if($cells[0].tagName -eq "TH") {
            $titles = @($cells | % { ("" + $_.InnerText).Trim() })
            continue
        }

        ## If we haven't found any table headers, make up names "P1", "P2", etc.
        if(-not $titles) {
            $titles = @(1..($cells.Count + 2) | % { "P$_" })
        }

        ## Now go through the cells in the the row. For each, try to find the
        ## title that represents that column and create a hashtable mapping those
        ## titles to content
        $resultObject = [Ordered] @{}
        for($counter = 0; $counter -lt $cells.Count; $counter++) {
            $title = $titles[$counter]
            if(-not $title) { continue }
            $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
        }

        ## And finally cast that hashtable to a PSCustomObject
        [PSCustomObject] $resultObject
    }
}#END: function Get-WebRequestTable

function Get-C2rSupportedVersions {
    [CmdletBinding()]
    param (
        # Major version (15 is ProPlus 2013, 16 is MS 365)
        [Parameter(Position=0)]
        [ValidateNotNull()]
        [ValidateSet(15,16)]
        [int]
        $MajorVersion=16
    )
            
    # URL for the 2013/ProPlus online chart we will scrape
    $Uri15page = 'https://docs.microsoft.com/en-us/officeupdates/update-history-office-2013'
    # this is going to be legacy at some point, but we will support it on "day 2".

    # URL for the MS 365 online chart we will scrape
    $Uri16page = 'https://docs.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date'

    # Determine the use case and set local variable for web request
    switch ($MajorVersion) {
        15 {
            $Uri = $Uri15page
        }
        16 {
            $Uri = $Uri16page
        }
    }

    # Get Web Request
    $WebRequest = Invoke-WebRequest -Uri $Uri
    $Table = Get-WebRequestTable $WebRequest -TableNumber 0

    # Determine the use case and convert data to object data types
    switch ($MajorVersion) {
        15 {
            foreach ($record in $Table) {

                # Release year is an int
                #   the year is sometime null, so we infer the previous value
                $year = if ($record.'Release year') {$record.'Release year' -as [int]} else {$year -as [int]}
                $record.'Release year' = $year
                # Release date is a partial date string. Leave as is for now
                # Version number is a version
                $record.'Version number' = $record.'Version number' -as [version]
                # More information is a link, leave as string but remove the space
                $record.'More information' = $record.'More information' -replace '\s'
                
                # Create a new property as a datetime based off two previous fields
                #   the date is a string we can split into a month and a day
                $tempArr = $record.'Release date' -split '\s'
                $month = $tempArr[0]
                $day = $tempArr[1]
                #  We can parse the string with its known format, then add the new member
                $ReleasedOn = [datetime]::parseexact("$($month)-$($day)-$($year)", 'MMMM-d-yyyy', $null)
                $record | Add-Member -MemberType NoteProperty -Name ReleasedOn -Value $ReleasedOn

            }#END: foreach ($record in $Table)
        }#END: 15
        16 {

            foreach ($record in $Table) {
                # Channel is a string, OK
                # Version is a number but sometimes is has a letter (20H2)
                # Build is a PART of a version number, but we can leave it as a string here
                # Release date is a datetime
                $record.'Release date' = $record.'Release date' -as [datetime]
                #Version supported until is a date but sometimes it's not. Leave as string for now
            }#END: foreach ($record in $Table)

            # Calculate which item should have latest build
            #  this will add a boolean property to Table
            $grpChannelSupported = $Table |
                Group-Object Channel
            $Table = Foreach ($G in $grpChannelSupported) {
                
                # If multiple items for this channel, do some logic
                $LatestBuildShouldBe = if (@($G.Group).Count -gt 1) {

                    # Grab the builds and cast as versions, sort and select
                    [string]$LastestBuild = $G.Group.Build |
                        Foreach-Object {[version]$_} |
                        Sort-Object -Descending |
                        Select-Object -First 1

                    # Choose the item with the latest build
                    $G.Group | Where-Object {$_.Build -eq $LastestBuild} |
                        Select-Object -Expand 'Build'
                    
                } else {

                    # There is only 1, choose it.
                    @($G.Group.Build)[0]

                }#END: $LatestBuildShouldBe = if (@($G.Group).Count -gt 1)

                # Tag the item with latest build
                $G.Group | Select-Object *, @{Name='isLatestBuild';Exp={
                    if ($_.Build -eq $LatestBuildShouldBe) {$true}else{$false}
                }}

            }#END: Foreach ($G in $grpChannelSupported)

        }#END: 16

    }#END: switch ($MajorVersion)


    Write-Output $Table

}#END: function Get-C2rSupportedVersions

function Get-C2rChannelInfo {
    <#
    .SYNOPSIS
    Gives the 'change' parameter value, GUID, and channel name for the given channel.
    .DESCRIPTION
    Returns the 'change' parameter value, GUID, and channel name required when changing the C2R update channel from the command line.
    .EXAMPLE
    $pValue = Get-C2rChannelInfo -ChannelName 'Monthly Enterprise Channel' | % ChangeParameterValue
    icm ('"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" /changesetting Channel={0}' -f $pValue)
    #>
    [CmdletBinding(DefaultParameterSetName='All')]
    param (
        # Search by Channel Name (Default)
        [Parameter(ParameterSetName='byChannelName')]
        [AllowNull()]
        [ValidateSet(
            'Current Channel',
            'Current (Preview)',
            'Semi-Annual Enterprise Channel',
            'Semi-Annual Enterprise Channel (Preview)',
            'Monthly Enterprise Channel',
            'Beta Channel'
        )]
        [string]
        $ChannelName,

        # Search by GUID
        [Parameter(ParameterSetName='byGuid')]
        [AllowNull()]
        [ValidateSet(
            '492350f6-3a01-4f97-b9c0-c7c6ddf67d60',
            '64256afe-f5d9-4f86-8936-8840a6a4f5be',
            '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114',
            'b8f9b850-328d-4355-9145-c59439a0c4cf',
            '55336b82-a18d-4dd6-b5f6-9e5095c314a6',
            '5440fd1f-7ecb-4221-8110-145efaa6372f',
            'f2e724c1-748f-4b47-8fb8-8e0d210e9208',
            '2e148de9-61c8-4051-b103-4af54baffbb4'
        )]
        [guid]
        $Guid,

        # Placeholder for null case
        [Parameter(ParameterSetName='All')]
        $All

    )

    # Define the table of objects in code
    $srcTable = @(
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
            ChangeParameterValue = 'Current'
            OfficialName         = 'Current Channel'
        },
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'64256afe-f5d9-4f86-8936-8840a6a4f5be'
            ChangeParameterValue = 'FirstReleaseCurrent'
            OfficialName         = 'Current (Preview)'
        },
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
            ChangeParameterValue = 'Broad'
            OfficialName         = 'Semi-Annual Enterprise Channel'
        },
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'b8f9b850-328d-4355-9145-c59439a0c4cf'
            ChangeParameterValue = 'Targeted'
            OfficialName         = 'Semi-Annual Enterprise Channel (Preview)'
        },
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'55336b82-a18d-4dd6-b5f6-9e5095c314a6'
            ChangeParameterValue = 'MonthlyEnterpise'
            OfficialName         = 'Monthly Enterprise Channel'
        },
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'5440fd1f-7ecb-4221-8110-145efaa6372f'
            ChangeParameterValue = 'BetaChannel'
            OfficialName         = 'Beta Channel'
        },
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'f2e724c1-748f-4b47-8fb8-8e0d210e9208'
            ChangeParameterValue = 'N/A'
            OfficialName         = 'N/A'
        },
        [pscustomobject]@{
            CdnUrlGuid           = [guid]'2e148de9-61c8-4051-b103-4af54baffbb4'
            ChangeParameterValue = 'N/A'
            OfficialName         = 'N/A'
        }
    )#END: $srcTable = @()
    
    Write-Debug "Parameter Set: $($PSCmdlet.ParameterSetName)"
    switch($PSCmdlet.ParameterSetName) {

        'byChannelName' {            
            $srcTable | Where-Object {$_.OfficialName -eq $ChannelName}
        }
        'byGuid'        {
            $srcTable | Where-Object {$_.CdnUrlGuid -eq $Guid}
        }
        'All'         {
            $srcTable
        }

    }#END: switch($PSCmdlet.ParameterSetName) {}
    
}#END: function Get-C2rChannelInfo

function Test-Ms365RequiresUpdate {
    <#
    .SYNOPSIS
    Checks the version info to see if the PC requires an update.
    .DESCRIPTION
    Checks the version info against the current online list to see if the item requires an update.
    .EXAMPLE
    Get-SoftwareList -Company XYZ -IncludeAppName 'Microsoft 365*' |
        Where {$_.ComputerName -eq 'PC1'} |
        Test-Ms365RequiresUpdate -Channel 'Monthly Enterprise Channel' |
        Select -Expand RequiresUpdate
    False
    False
    .NOTES
    v1.0 will support basic check of version vs channel name
    Later versions will include 
    - check vs channel guid
    - Option to pass if version is not latest but is still supported.
    #>
    [CmdletBinding(DefaultParameterSetName='byChannelName')]
    param (
        # Given Channel NAME to check the version against
        [Parameter(Mandatory=$true,
            ParameterSetName='byChannelName')]
        [string]
        $Channel,

        # Given Channel GUID to check the version against
        [Parameter(Mandatory=$true,
            ParameterSetName='byUniqueID')]
        [guid]
        $Guid,

        # Object containing Version and ComputerName from the PC to test. Must not include non-MS365 software items.
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true)]
        [ValidateScript({
            ([version]$_.Version) -is [version] -and
            $_.ComputerName -is [string]
        })]
        [PsCustomObject]
        $InputObject
    )
    
    begin {
        
        # Cache the current list of supported versions
        $C2rSupportedVersions = Get-C2rSupportedVersions -MajorVersion 16

        # Define an output object
        $OutputObject = New-Object System.Collections.ArrayList 

    }
    
    process {
        
        Foreach ($obj in $InputObject) {

            # Find the item for comparison
            $LatestChannelBuild = $C2rSupportedVersions | Where-Object {
                $_.isLatestBuild -and
                $_.channel -eq $channel                
            } | Select-Object -Expand 'Build'

            # Does the computer have the latest version?
            $objBuild = "$(([version]($obj.Version)).Build).$(([version]($obj.Version)).Revision)"
            $isLatestVersion = [version]($LatestChannelBuild) -le ([version]$objBuild)

            # Create an output object with names, version, boolean
            $thisObj = [PSCustomObject]@{
                ComputerName = $obj.ComputerName
                Channel = $Channel
                RequiresUpdate = !$isLatestVersion
                BuildShouldBe = $LatestChannelBuild
                BuildIs = $objBuild
                isLatestVersion = $isLatestVersion
                ComputerId = $obj.ComputerId
            }

            [void]($OutputObject.Add($thisObj))
        }
    }
    
    end {
        Write-Output $OutputObject
    }

}#END: function Test-Ms365RequiresUpdate

# Use available handshake protcols
[Net.ServicePointManager]::SecurityProtocol = 
[enum]::GetNames([Net.SecurityProtocolType]) | Foreach-Object {
    [Net.SecurityProtocolType]::$_
}
