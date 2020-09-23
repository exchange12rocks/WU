function Get-LatestWindowsRollup {
    <# MIT License

    Copyright (c) 2017 Kirill Nikolaev

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
    #>
    
    
    <#
    .SYNOPSIS
    Retrieves the last available rollup update ID from the Microsoft Knowledge Base.
    
    .DESCRIPTION
    Retrieves the last available rollup update ID from the Microsoft Knowledge Base.
    
    .PARAMETER OS
    Operating System version, for which to retrieve the latest rollup. Examples: 2012R2, 1703, 2019
    
    .EXAMPLE
    Get-LatestWindowsRollup -OS 2012R2
    
    .NOTES
    Author: Kirill Nikolaev
    Twitter: @exchange12rocks
    Web-site: https://exchange12rocks.org
    GitHub: https://github.com/exchange12rocks
    
    .LINK
    https://exchange12rocks.org
    
    .LINK
    https://github.com/exchange12rocks/WU/tree/master/Get-LatestWindowsRollup
    
    #>

    #Requires -Version 3.0
    
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [ValidateSet('2012', '2012R2', '2016', '1607', '1703', '1709', '1803', '1809', '2019')]
        [string]$OS,
        [ValidateSet('All', 'RollupsOnly')]
        [string]$Mode = 'RollupsOnly'
    )
    
    $ErrorActionPreference = 'Stop'
    
    try {
        $SupportKBUriTemplate = 'https://support.microsoft.com/app/content/api/content/help/en-us/{0}'
        $PreviewRollupRegEx = ('[A-Za-z]+ \d{{1,2}}, \d{{4}}[{0}-]KB\d+ \(Preview of Monthly Rollup\)' -f [char](8212)) # Somehow PS5.1 cannot process the "—" character in strings correctly.
        $SecurityOnlyRegEx = ('[A-Za-z]+ \d{{1,2}}, \d{{4}}[{0}-]KB\d+ \(Security-only update\)' -f [char](8212))
        $PreviewRegEx = ('\w+ \d{{1,2}}, \d{{4}}[{0}-]KB\d+ \(\w{{2}} \w{{5}} \d+\.\d+\) Preview$' -f [char](8212))
        $OSDefs = @{
            '2012'   = @{
                KB                    = '4009471'
                SeveralTypesAvailable = $true
            }
            '2012R2' = @{
                KB                    = '4009470'
                SeveralTypesAvailable = $true
            }
            '2016'   = @{
                KB = '4000825'
            }
            '1607'   = @{
                KB = '4000825'
            }
            '1703'   = @{
                KB = '4018124'
            }
            '1709'   = @{
                KB = '4043454'
            }
            '1803'   = @{
                KB = '4099479'
            }
            '1809'   = @{
                KB = '4464619'
            }
            '2019'   = @{
                KB = '4464619'
                SeveralTypesAvailable = $true
            }
        }

        $OSSupportKBUri = $SupportKBUriTemplate -f $OSDefs.$OS.KB
        $OSSupportKB = Invoke-WebRequest -Uri $OSSupportKBUri -UseBasicParsing
        $OSSupportKBContent = $OSSupportKB.Content
        $OSSupportKBContentConverted = ConvertFrom-Json -InputObject $OSSupportKBContent
        if ($OSSupportKBContentConverted.releaseNoteRelationship.minorVersions) {
            # If there were any updates for this version released, this property should be populated. If it is unpopulated - it means there are no updates.
            $UpdateList = $OSSupportKBContentConverted.releaseNoteRelationship.minorVersions
            $UpdateListProcessed = [System.Array]::CreateInstance([PSCustomObject], $UpdateList.Count)
            for ($Counter = 0; $Counter -le ($UpdateList.Count - 1); $Counter++) {
                $UpdateRecord = $UpdateList[$Counter]
                $UpdateRecordDate = Get-Date -Date $UpdateRecord.releaseDate
                $UpdateListProcessed[$Counter] = [pscustomobject]@{
                    ID    = $UpdateRecord.id
                    Date  = $UpdateRecordDate
                    Title = $UpdateRecord.heading
                }
            }
            $UpdateListProcessedSorted = Sort-Object -InputObject $UpdateListProcessed -Property 'Date' -Descending

            if ($OSDefs.$OS.SeveralTypesAvailable) {
                switch ($Mode) {
                    'All' {
                        $RollupID = $null
                        $RollupPreviewID = $null
                        $SecurityOnlyID = $null
                        foreach ($UpdateRecordProcessed in $UpdateListProcessedSorted) {
                            if ($UpdateRecordProcessed.Title -match $SecurityOnlyRegEx) {
                                if (-not $SecurityOnlyID) {
                                    $SecurityOnlyID = $UpdateRecordProcessed.ID
                                }
                            }
                            elseif ($UpdateRecordProcessed.Title -match $PreviewRollupRegEx -or $UpdateRecordProcessed.Title -match $PreviewRegEx) {
                                if (-not $RollupPreviewID) {
                                    $RollupPreviewID = $UpdateRecordProcessed.ID
                                }
                            }
                            elseif (-not $RollupID) {
                                $RollupID = $UpdateRecordProcessed.ID
                            }

                            if ($RollupID -and $RollupPreviewID -and $SecurityOnlyID) {
                                break
                            }
                        }

                        @{
                            RollupID        = $RollupID
                            RollupPreviewID = $RollupPreviewID
                            SecurityOnlyID  = $SecurityOnlyID
                        }
                    }
                    Default {
                        foreach ($UpdateRecordProcessed in $UpdateListProcessedSorted) {
                            if ($UpdateRecordProcessed.Title -notmatch $PreviewRollupRegEx -and $UpdateRecordProcessed.Title -notmatch $SecurityOnlyRegEx -and $UpdateRecordProcessed.Title -notmatch $PreviewRegEx) {
                                $UpdateRecordProcessed.ID
                                break
                            }
                        }
                    }
                }
            }
            else {
                $UpdateListProcessedSorted[0].ID
            }
        }
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}