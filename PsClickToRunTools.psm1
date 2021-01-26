function Get-SupportedVersions {
    [CmdletBinding()]
    param (
        # Major version (15 is ProPlus 2013, 16 is MS 365)
        [Parameter(Position=0)]
        [ValidateNotNull()]
        [int]
        $MajorVersion=16
    )
            
    # URL for the 2013/ProPlus online chart we will scrape
    $Uri15page = ''
    $Uri15title = ''

    # URL for the MS 365 online chart we will scrape
    $Uri16page = 'https://docs.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date'
    $Uri16title = 'Supported Versions'

    # 

}#END: function Get-SupportedVersions
