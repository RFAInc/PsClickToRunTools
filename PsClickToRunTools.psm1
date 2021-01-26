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

    # Determine the use case and set local variable for web request
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
                # Version is an int
                $record.Version = $record.Version -as [int]
                # Build is a part of a version number, but we can leave it as a string here
                # Release date is a datetime
                $record.'Release date' = $record.'Release date' -as [datetime]
                #Version supported until is a date but sometimes it's not. Leave as string for now
            }#END: foreach ($record in $Table)

        }#END: 16
    }#END: switch ($MajorVersion)


    Write-Output $Table

}#END: function Get-C2rSupportedVersions
