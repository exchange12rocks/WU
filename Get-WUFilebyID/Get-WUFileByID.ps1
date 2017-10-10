function Get-WUFileByID {

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
Retrieves update files from Microsoft Update Catalog.

.DESCRIPTION
Retrieves update files from Microsoft Update Catalog by a KB number or a GUID of an update set.

.PARAMETER KB
Knowledge base article number. Example: 4038793

.PARAMETER GUID
GUID of an update set. The GUID uniquely identifies a set of files for a KB article-product pair.

.PARAMETER SearchCriteria
Product name (or a part of it). If an update has been released for several products (like different OSes), use the parameter to specify which product ot target.

.PARAMETER DestinationFolder
A folder where downloaded files will be saved.

.PARAMETER LinksOnly
Instructs the function to return only links to files w/o downloading them.

.PARAMETER ForceSSL
By default the Update Catalog returns HTTP links. Use this parameter to download files through HTTPS.
Note that due to differences in CDN nodes set up, sometimes HTTPS downloads may fail.

.PARAMETER SearchPageTemplate
A URL of a web-page with a list of available update sets for a KB.

.PARAMETER DownloadPageTemplate
A URL of a web page with a list of available files for an update set.

.EXAMPLE
# Download files for KB article #3172729 for Windows Server 2012 R2.
Get-WUFileByID -KD 3172729 -SearchCriteria 'Windows Server 2012 R2'

.EXAMPLE
# Download update files by a known update set GUID
Get-WUFileByID -GUID 2c61a788-27e5-44f9-b27b-1ca22b4592d9

.EXAMPLE
# Download all updates required for a computer

$Criteria = "IsInstalled=0 and Type='Software'"
$Searcher = New-Object -ComObject Microsoft.Update.Searcher
$SearchResult = $Searcher.Search($Criteria).Updates
[string[]]$KBIDs = ''
($SearchResult | Select-Object Title).Title | ForEach-Object{
    $null = $_ -match '^.+\(KB(\d+)\)$'
    $KBIDs += $matches[1]
}
foreach ($ID in $KBIDs) {
    Get-WUFileByID -KD $ID -SearchCriteria 'Windows Server 2008 R2'
}

.NOTES
Author: Kirill Nikolaev
Twitter: @exchange12rocks
Web-site: https://exchange12rocks.org
GitHub: https://github.com/exchange12rocks

.LINK
https://exchange12rocks.org/2017/10/02/function-to-download-updates-from-microsoft-catalog

.LINK
https://github.com/exchange12rocks/PS/tree/master/Get-WUFilebyID

#>

    #Requires -Version 3.0

    Param (
        [Parameter(ParameterSetName = 'Default', Position = 0, Mandatory)]
        [string]$KB,
        [Parameter(ParameterSetName = 'ByGUID', Position = 0, Mandatory)]
        [guid]$GUID,
        [Parameter(ParameterSetName = 'Default', Position = 1, Mandatory)]
        [string]$SearchCriteria,
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'ByGUID')]
        [ValidateScript({Test-Path $_ -PathType 'Container'})]
        [string]$DestinationFolder = '.\',
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'ByGUID')]
        [switch]$LinksOnly,
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'ByGUID')]
        [switch]$ForceSSL,
        [Parameter(ParameterSetName = 'Default')]
        [string]$SearchPageTemplate = 'https://www.catalog.update.microsoft.com/Search.aspx?q={0}',
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'ByGUID')]
        [string]$DownloadPageTemplate = 'https://www.catalog.update.microsoft.com/DownloadDialog.aspx?updateIDs=%5B%7B%22size%22%3A0%2C%22languages%22%3A%22%22%2C%22uidInfo%22%3A%22{0}%22%2C%22updateID%22%3A%22{0}%22%7D%5D&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version='
    )

    BEGIN {

        $ErrorActionPreference = 'Stop'

        if ($PSCmdlet.ParameterSetName -eq 'Default') {
            function FindTableColumnNumber {
                Param (
                    [Parameter(Mandatory)]
                    $Columns,
                    [Parameter(Mandatory)]
                    [string]$Pattern
                )
    
                BEGIN {
                    $Counter = 0
                }
                PROCESS {
                    foreach ($Column in $Columns) {
                        if ($Column.InnerHTML -like $Pattern) {
                            break
                        }
                        $Counter++
                    }
                }
                END {
                    return $Counter
                }
            }
        }

        function DownloadWUFile {
            Param (
                [Parameter(Mandatory)]
                [string]$URL,
                [Parameter(Mandatory)]
                [ValidateScript({Test-Path $_ -PathType 'Container'})]
                [string]$DestinationFolder,
                [Parameter(Mandatory)]
                [string]$FileName,
                [string]$KB,
                [string]$GUID
            )

            BEGIN {
                $Result = @()
            }
            PROCESS {
                $FullFileName = Join-Path -Path $DestinationFolder -ChildPath $FileName
                Invoke-WebRequest -Uri $URL -OutFile $FullFileName
                if (Test-Path -Path $FullFileName -PathType Leaf) {
                    $Result += (Get-Item -Path $FullFileName)
                }
                else {
                    if ($GUID) {
                        Write-Error -Message ('Could not write a file for GUID {0} to a location {1}' -f $GUID, $FullFileName)
                    }
                    else {
                        Write-Error -Message ('Could not write a file for KB {0} to a location {1}' -f $KB, $FullFileName)
                    }
                }
            }
            END {
                return $Result
            }
        }

        function RewriteURLtoHTTPS {
            Param (
                [Parameter(Mandatory)]
                [string[]]$URL
            )

            BEGIN {
                $Result = @()
            }
            PROCESS {
                foreach ($Item in $URL) {
                    if ($Item -match '^https*:\/\/(.+)$') {
                        $Result += ('https://{0}' -f $Matches[1])
                    }
                }
            }
            END {
                return $Result
            }
        }

        function ParseKBDownloadLinksFromText {
            Param (
                [Parameter(Mandatory)]
                $Text,
                [guid]$GUID,
                [string]$KB,
                [switch]$ForceSSL
            )

            BEGIN {
                $Result = @()
            }
            PROCESS {
                foreach ($Item in ((Select-String -InputObject $Text -Pattern "downloadInformation\[\d\]\.files\[\d\]\.url\s*=\s*'.+'\s*;" -AllMatches).Matches)) { # There could be several files for a single update - gotta catch them all!
                    if ($Item -match "downloadInformation\[\d\]\.files\[\d\]\.url\s*=\s*'(.+)'\s*;") {
                        $Result += $Matches[1]
                    }
                    else {
                        if ($KB) {
                            Write-Error -Message ('Could not extract a download link from the element {0} for KB {1}' -f $Item, $KB)
                        }
                        else {
                            Write-Error -Message ('Could not extract a download link from the element {0} for GUID {1}' -f $Item, $GUID)
                        }
                    }
                }
            }
            END {
                if ($ForceSSL) {
                    $Result = RewriteURLtoHTTPS -URL $Result
                }
                return $Result
            }
        }
        function ParseKBDownloadLinksFromHTML {
            Param (
                [Parameter(Mandatory)]
                [Microsoft.PowerShell.Commands.HtmlWebResponseObject]$HTMLObject,
                [guid]$GUID,
                [string]$KB,
                [switch]$ForceSSL
            )

            BEGIN {
            }
            PROCESS {
                $KBCatalogDownloadPageScripts = $HTMLObject.ParsedHtml.getElementsByTagName('script') # The actual update links are in a JavaScript array, grabbing all the scripts from the page. 
                $KBDownloadScriptText = ($KBCatalogDownloadPageScripts | Where-Object -FilterScript {$_.innerHTML -Like '*var downloadInformation = new Array();*'}).innerHTML # Then we finding the one containing the download links.
                $Result = ParseKBDownloadLinksFromText -Text $KBDownloadScriptText -KB $KB -GUID $GUID -ForceSSL:$ForceSSL
            }
            END {
                return $Result
            }
        }

        function GetKBDownloadLinksByGUID {
            Param (
                [Parameter(Mandatory)]
                [string]$DownloadPageTemplate,
                [Parameter(Mandatory)]
                [guid]$GUID,
                [string]$KB,
                [switch]$ForceSSL
            )

            BEGIN {
            }
            PROCESS {
                $KBCatalogDownloadPage = Invoke-WebRequest -Uri ($DownloadPageTemplate -f $GUID)
                $Result = ParseKBDownloadLinksFromHTML -HTMLObject $KBCatalogDownloadPage -KB $KB -GUID $GUID -ForceSSL:$ForceSSL
            }
            END {
                return $Result
            }
        }
    }
    PROCESS {
        if ($PSCmdlet.ParameterSetName -eq 'Default') {
            $KBCatalogPage = Invoke-WebRequest -Uri ($SearchPageTemplate -f $KB)
            $Rows = $KBCatalogPage.ParsedHtml.getElementById('ctl00_catalogBody_updateMatches').getElementsByTagName('tr') # This line detects the main table which contains updates data.
            
            $HeaderRow = $null
            foreach ($Row in $Rows) {
                if ($Row.id -eq 'headerRow') {
                    if (-not $HeaderRow) {
                        $HeaderRow = $Row
                    }
                    else {
                        Write-Error -Message ('Multiple header rows returned for KB {0}' -f $KB)
                    }
                }
            }

            if ($HeaderRow) {
                $Columns = $HeaderRow.getElementsByTagName('td')
            }
            else {
                Write-Error -Message ('Could not find a header row for KB {0}' -f $KB)
            }

            $DateColumnNumber = FindTableColumnNumber -Columns $Columns -Pattern '*<SPAN>Last Updated</SPAN>*' # Finding a column where the update release date is stored.
            $ProductColumnNumber = FindTableColumnNumber -Columns $Columns -Pattern '*<SPAN>Products</SPAN>*' # Finding a column where the product name is stored.
            $TitleColumnNumber = FindTableColumnNumber -Columns $Columns -Pattern '*<SPAN>Title</SPAN>*' # Finding a column where update title and ID are stored.

            $DataRows = @()
            foreach ($Row in $Rows) {
                if ($Row.id -ne 'headerRow') {
                    $DataRows += $Row
                }
            }
            if (-not $DataRows) {
                Write-Error -Message ('No data rows have been found for KB {0}' -f $KB)
            }

            $CandidateRows = @()
            foreach ($Row in $DataRows) { # There could be not one but several rows matching the pattern. We should process them all. $CandidateRows contains such rows.
                if ($Row.getElementsByTagName('td')[$ProductColumnNumber].innerHTML -like ('*{0}*' -f $SearchCriteria)) {
                    $CandidateRows += $Row
                }
            }
            if (-not $CandidateRows) {
                Write-Error -Message ('No candidate rows have been found for KB {0}' -f $KB)
            }

            $RowNumber = 0
            $ReleaseDate = New-Object -TypeName DateTime -ArgumentList @(1, 1, 1) # The lowest possible date in .NET
            foreach ($Row in $CandidateRows) { # Here we are looking for a row with the most recent release date.
                if ($ReleaseDate -lt [datetime]::ParseExact($Row.getElementsByTagName('td')[$DateColumnNumber].innerHTML.Trim(), 'm/d/yyyy', $null)) { # We assume that MS never publishes several versions of an update on the same day.
                    break
                }
                $RowNumber++
            }

            if ($CandidateRows[$RowNumber].getElementsByTagName('td')[$TitleColumnNumber].innerHTML -match 'goToDetails\("(.+)"\);') { # goToDetails contains update's GUID which we use then to request an update download page.
                $GUID = $matches[1]
            }
            else {
                Write-Error -Message ('Could not find a GUID for KB {0}' -f $KB)
            }
        }

        if ($GUID -is [Guid]) { # Since the function parameter should be of the [guid] type, to simplify validation, and we actually need a string (and set $GUID to a string if the work mode is NOT "ByGUID"), it's easier to just check&convert the type on-the-fly.
            [string]$GUID = $Guid.Guid
        }

        if ($PSCmdlet.ParameterSetName -eq 'Default') {
            $DownloadLinks = GetKBDownloadLinksByGUID -DownloadPageTemplate $DownloadPageTemplate -GUID $GUID -KB $KB -ForceSSL:$ForceSSL
        }
        else {
            $DownloadLinks = GetKBDownloadLinksByGUID -DownloadPageTemplate $DownloadPageTemplate -GUID $GUID -ForceSSL:$ForceSSL
        }
        
        if ($DownloadLinks) {
            if ($LinksOnly) {
                $Result = $DownloadLinks
            }
            else {
                $Result = @()
                foreach ($URL in $DownloadLinks) { 
                    if ($URL -match '.+/(.+)$') {
                        if ($PSCmdlet.ParameterSetName -eq 'Default') {
                            $Result += DownloadWUFile -URL $URL -DestinationFolder $DestinationFolder -FileName $Matches[1] -KB $KB
                        }
                        else {
                            $Result += DownloadWUFile -URL $URL -DestinationFolder $DestinationFolder -FileName $Matches[1] -GUID $GUID
                        }
                    }
                    else {
                        if ($PSCmdlet.ParameterSetName -eq 'Default') {
                            Write-Error -Message ('URL {0} for KB {1} does not match the scheme' -f $URL, $KB)
                        }
                        else {
                            Write-Error -Message ('URL {0} for GUID {1} does not match the scheme' -f $URL, $GUID)
                        }
                    }
                }
            }
        }
        else {
            if ($PSCmdlet.ParameterSetName -eq 'Default') {
                Write-Error -Message ('No download links for KB {0} have been generated' -f $KB)
            }
            else {
                Write-Error -Message ('No download links for GUID {0} have been generated' -f $GUID)
            }
        }
    }
    END {
        return $Result
    }
}