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
Operating System version, for which to retrieve the latest rollup. Examples: 2012R2, 10-1703

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

    Param (
        [Parameter(Mandatory)]
        [ValidateSet('2008R2','2012','2012R2','2016','1703','1709')]
        [string]$OS,
        [ValidateSet('All','RollupsOnly','NonRollupsOnly')]
        [string]$Mode = 'All'
    )

    $ErrorActionPreference = 'Stop'

    $SupportKBUriTemplate = 'https://support.microsoft.com/app/content/api/content/help/en-us/{0}'
    $OSDefs = @{
        '2008R2' = @{
            KB = '4009469'
            RegEx = '\d{4}\-\d{2}\sMonthly\sRollup\s\-\sKB(?<KB>\d+)'
            KBRegEx = '\d{4}\-\d{2}.*\sKB(?<KB>\d+)'
            ParseDate = $true
        }
        '2012' = @{
            KB = '4009471'
            RegEx = '\d{4}\-\d{2}\sMonthly\sRollup\s\-\sKB(?<KB>\d+)'
            KBRegEx = '\d{4}\-\d{2}.*\sKB(?<KB>\d+)'
            ParseDate = $true
        }
        '2012R2' = @{
            KB = '4009470'
            RegEx = '\d{4}\-\d{2}\sMonthly\sRollup\s\-\sKB(?<KB>\d+)'
            KBRegEx = '\d{4}\-\d{2}.*\sKB(?<KB>\d+)'
            ParseDate = $true
        }
        '2016' = @{
            KB = '4000825'
            RegEx = 'KB(?<KB>\d+)\s\(OS\sBuild\s\d+\.\d+\)'
            ParseDate = $false
            DisplayNameFilter = 'Windows 10 Version 1607 and Windows Server 2016'
        }
        '1703' = @{
            KB = '4018124'
            RegEx = 'KB(?<KB>\d+)\s\(OS\sBuild\s\d+\.\d+\)'
            ParseDate = $false
            DisplayNameFilter = 'Windows 10 Version 1703'
        }
        '1709' = @{
            KB = '4043454'
            RegEx = 'KB(?<KB>\d+)\s\(OS\sBuild\s\d+\.\d+\)'
            ParseDate = $false
            DisplayNameFilter = 'Windows 10 Version 1709'
        }
    }

    $ToSort = @()
    $UpdatesList = ((Invoke-WebRequest -Uri ($SupportKBUriTemplate -f (((Invoke-WebRequest -Uri ($SupportKBUriTemplate -f $OSDefs.$OS.KB) -UseBasicParsing).Content | ConvertFrom-Json).sideNav)) -UseBasicParsing).Content | ConvertFrom-Json).links
    if ($OSDefs.$OS.ParseDate) {
        $FilteredUpdatesList = @()
        foreach ($Item in $UpdatesList) {
            switch ($Mode) {
                'RollupsOnly' {
                    if ($Item.text -match $OSDefs.$OS.RegEx) {
                        $FilteredUpdatesList += $Item
                    }
                }
                'NonRollupsOnly' {
                    if ($Item.text -notmatch $OSDefs.$OS.RegEx) {
                        $FilteredUpdatesList += $Item
                    }
                }
            }
        }

        if ($FilteredUpdatesList.Count -eq 0) {
            $FilteredUpdatesList = $UpdatesList
        }

        $DateRegEx = '(?<Year>\d{4})\-(?<Month>\d{2}).+'
        foreach ($Item in $FilteredUpdatesList) {
            if ($Item.text -match $DateRegEx) {
                $ToSort += [pscustomobject]@{
                    ID = $Item.ID
                    Year = $Matches.Year
                    Month = $Matches.Month
                }
            }
        }

        if ($ToSort) {
            $YearOfInterest = ($ToSort.Year | Select-Object -Unique | Sort-Object -Descending)[0]
            if ($YearOfInterest) {
                $MonthOfInterest = (($ToSort | Where-Object -Property Year -EQ -Value $YearOfInterest).Month | Select-Object -Unique | Sort-Object -Descending)[0]
                if ($MonthOfInterest) {
                    $IDsOfInterest = ($ToSort | Where-Object -FilterScript {$_.Year -eq $YearOfInterest -and $_.Month -eq $MonthOfInterest}).ID

                    if ($IDsOfInterest) {
                        $TextOfInterest = @()

                        foreach ($Item in $UpdatesList) {
                            if ($Item.id -in $IDsOfInterest) {
                                $TextOfInterest += $Item.text
                            }
                        }

                        if ($TextOfInterest) {
                            $Result = @()
                            foreach ($Item in $TextOfInterest) {
                                if ($Item -match $OSDefs.$OS.KBRegEx) {
                                    $Result += $Matches.KB
                                }
                            }
                            return $Result
                        }
                        else {
                            Write-Error -Message ('Could not find updates by ID #s {0}' -f $IDsOfInterest)
                        }
                    }
                    else {
                        Write-Error -Message ('Could not find updates for the year {0} and the month {1}' -f $YearOfInterest, $MonthOfInterest)
                    }
                }
                else {
                    Write-Error -Message 'Could not extract a last month number from update titles'
                }
            }
            else {
                Write-Error -Message 'Could not extract a last year number from update titles'
            }
        }
        else {
            Write-Error -Message ('No items matched regular expression {0}' -f $RegEx)
        }
    }
    else {
        for ($Counter=0;($UpdatesList.Count-1);$Counter++) {
            if ($UpdatesList[$Counter].text -eq $OSDefs.$OS.DisplayNameFilter) {
                if ($UpdatesList[$Counter+1].text -match $OSDefs.$OS.RegEx) {
                    return $Matches.KB
                }
                else {
                    Write-Error -Message ('Expected update entry {0} does not match regular expression {1}' -f $UpdatesList[$Counter+1].text, $OSDefs.$OS.RegEx) 
                }
            }
        }

        Write-Error -Message ('Could not find title "{0}" in the list' -f $OSDefs.$OS.DisplayNameFilter)
    }
}