#Requires -Version 3.0
#Requires -Modules Pester

$FunctionName = ($MyInvocation.MyCommand.Name).Substring(0, ($MyInvocation.MyCommand.Name).Length - 10)
. (Join-Path -Path (Split-Path -Path $MyInvocation.MyCommand.Path) -ChildPath ('{0}.ps1' -f $FunctionName))

$SingleFileKBID = '4038793'
$SingleFileGUID = '9516efa8-6493-43de-979c-ebf2aa89aa69'
$SingleFileName = 'windows8.1-kb4038793-x64_14934122608b58b49cbe6fb3a8834010544f2263.msu'
$SingleFileNameRegEx = 'windows8\.1-kb4038793-x64_14934122608b58b49cbe6fb3a8834010544f2263\.msu'
$MultipleFilesKBID = '3172729'
$MultipleFilesGUID = 'cdde339c-ebdb-4a16-add4-fb196a5053a8'
$MultipleFilesNames = @(
    'windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu'
    'windows8.1-kb3172729-x64_e8003822a7ef4705cbb65623b72fd3cec73fe222.msu'
)
$MultipleFilesNamesRegEx = @(
    'windows8\.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c\.msu'
    'windows8\.1-kb3172729-x64_e8003822a7ef4705cbb65623b72fd3cec73fe222\.msu'
)
$SearchCriteria = 'Windows Server 2012 R2'

$MultipleFilesKBIDSC = '4014524'
$MultipleFilesGUIDSC = '2c61a788-27e5-44f9-b27b-1ca22b4592d9'
$MultipleFilesNamesSC = @(
    'kb4014524_adminconsole_i386_168bff08cf368995926bf6eacc1739ea216d2f25.cab',
    'kb4014524_adminconsole_amd64_eee3fbe82491d3b5c86eb3759cc6530b11893590.cab'
)
$MultipleFilesNamesSCRegEx = @(
    'kb4014524_adminconsole_i386_168bff08cf368995926bf6eacc1739ea216d2f25\.cab',
    'kb4014524_adminconsole_amd64_eee3fbe82491d3b5c86eb3759cc6530b11893590\.cab'
)
$SearchCriteriaSC = '*'

$HTTPMultiple = @(
    'http://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu',
    'http://download.windowsupdate.com/d/msdownload/update/software/secu/2016/07/windows8.1-kb3172729-x64_e8003822a7ef4705cbb65623b72fd3cec73fe222.msu'
)
$HTTPSMultiple = @(
    'https://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu',
    'https://download.windowsupdate.com/d/msdownload/update/software/secu/2016/07/windows8.1-kb3172729-x64_e8003822a7ef4705cbb65623b72fd3cec73fe222.msu'
)
$HTTPSingle = @(
    'http://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu'
)
$HTTPSSingle = @(
    'https://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu'
)
$HTTPSingleString = 'http://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu'
$HTTPSSingleString = 'https://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu'


Describe -Name 'Pester presence' -Fixture {
    $Module = Get-Module -Name 'Pester'
    It -name 'The module is loaded' -test {
        $Module | Should -Not -BeNullOrEmpty
    }

    It -name 'The module version is appropriate' -test {
        $Module.Version | Should -BeGreaterThan 4.0.7
    }
}

Describe -Name 'Full tests' -Fixture {
    Push-Location -Path 'TestDrive:\'
    Context -Name 'Single file - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $SingleFileKBID -SearchCriteria $SearchCriteria
        
        It -name 'All files are downloaded' -test {
            foreach ($Item in $SingleFileName) {
                $Item | Should -Exist
            }
        }
        It -name 'Only a single item returned' -test {
            $FunctionResult.Count | Should -Be 1
        }
        It -name 'Returned items are of System.IO.FileInfo type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.IO.FileInfo'
            }
        }

        foreach ($Item in $SingleFileName) {
            Remove-Item -Path $Item -Force
        }
    }
    Context -Name 'Multiple files - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $MultipleFilesKBID -SearchCriteria $SearchCriteria

        It -name 'All files are downloaded' -test {
            foreach ($Item in $MultipleFilesNames) {
                $Item | Should -Exist
            }
        }
        It -name 'Multiple items returned' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'Returned items are of System.IO.FileInfo type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.IO.FileInfo'
            }
        }

        foreach ($Item in $MultipleFilesNames) {
            Remove-Item -Path $Item -Force
        }
    }
    Context -Name 'Single file - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $SingleFileGUID
        
        It -name 'All files are downloaded' -test {
            foreach ($Item in $SingleFileName) {
                $Item | Should -Exist
            }
        }
        It -name 'Only a single item returned' -test {
            $FunctionResult.Count | Should -Be 1
        }
        It -name 'Returned items are of System.IO.FileInfo type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.IO.FileInfo'
            }
        }

        foreach ($Item in $SingleFileName) {
            Remove-Item -Path $Item -Force
        }
    }
    Context -Name 'Multiple files - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $MultipleFilesGUID
            
        It -name 'Multiple files downloaded by GUID' -test {
            foreach ($Item in $MultipleFilesNames) {
                $Item | Should -Exist
            }
        }
        It -name 'Multiple files returned by GUID' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'Returned items are of System.IO.FileInfo type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.IO.FileInfo'
            }
        }

        foreach ($Item in $MultipleFilesNames) {
            Remove-Item -Path $Item -Force
        }
    }
    Context -Name 'SC - Multiple files - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $MultipleFilesKBIDSC -SearchCriteria $SearchCriteriaSC

        It -name 'All files are downloaded' -test {
            foreach ($Item in $MultipleFilesNamesSC) {
                Test-Path -Path $Item | Should -Be $true
            }
        }
        It -name 'Multiple items returned' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'Returned items are of System.IO.FileInfo type' -test {
            foreach ($Item in $FunctionResult) {
                $Item -is [System.IO.FileInfo] | Should Be $true
            }
        }

        foreach ($Item in $MultipleFilesNamesSC) {
            Remove-Item -Path $Item -Force
        }
    }
    Context -Name 'SC - Multiple files - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $MultipleFilesGUIDSC
            
        It -name 'Multiple files downloaded by GUID' -test {
            foreach ($Item in $MultipleFilesNamesSC) {
                $Item | Should -Exist
            }
        }
        It -name 'Multiple files returned by GUID' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'Returned items are of System.IO.FileInfo type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.IO.FileInfo'
            }
        }

        foreach ($Item in $MultipleFilesNamesSc) {
            Remove-Item -Path $Item -Force
        }
    }
    Pop-Location
}

Describe -Name 'Links only' -Fixture {
    Context -Name 'Single file - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $SingleFileKBID -SearchCriteria $SearchCriteria -LinksOnly
        It -Name 'Only a single item returned' -test {
            $FunctionResult.Count | Should -Be 1
        }
        It -name 'The correct item is returned' -test {
            $FunctionResult | Should -Match ('^https*:\/\/.+{0}$' -f $SingleFileNameRegEx)
        }
        It -name 'Returned item is of System.String type' -test {
            $FunctionResult | Should -BeOfType 'System.String'
        }
    }
    Context -Name 'Multiple files - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $MultipleFilesKBID -SearchCriteria $SearchCriteria -LinksOnly
        It -Name 'Multiple items returned' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
    Context -Name 'Single file - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $SingleFileGUID -LinksOnly
        It -Name 'Only a single item returned' -test {
            $FunctionResult.Count | Should -Be 1
        }
        It -name 'The correct item is returned' -test {
            $FunctionResult | Should -Match ('^https*:\/\/.+{0}$' -f $SingleFileNameRegEx)
        }
        It -name 'Returned item is of System.String type' -test {
            $FunctionResult | Should -BeOfType 'System.String'
        }
    }
    Context -Name 'Multiple files - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $MultipleFilesGUID -LinksOnly
        It -name 'Multiple files returned by GUID' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
    Context -Name 'SC - Multiple files - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $MultipleFilesKBIDSC -SearchCriteria $SearchCriteriaSC -LinksOnly
        It -name 'Multiple items returned' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesSCRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesSCRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
    Context -Name 'SC - Multiple files - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $MultipleFilesGUIDSC -LinksOnly
        It -name 'Multiple files returned by GUID' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesSCRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesSCRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
}

Describe -Name 'Links only - HTTPS' -Fixture {
    Context -Name 'Single file - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $SingleFileKBID -SearchCriteria $SearchCriteria -LinksOnly -ForceSSL
        It -Name 'Only a single item returned' -test {
            $FunctionResult.Count | Should -Be 1
        }
        It -name 'The correct item is returned' -test {
            $FunctionResult | Should -Match ('^https*:\/\/.+{0}$' -f $SingleFileNameRegEx)
        }
        It -name 'Returned item is of System.String type' -test {
            $FunctionResult | Should -BeOfType 'System.String'
        }
    }
    Context -Name 'Multiple files - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $MultipleFilesKBID -SearchCriteria $SearchCriteria -LinksOnly -ForceSSL
        It -Name 'Multiple items returned' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
    Context -Name 'Single file - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $SingleFileGUID -LinksOnly -ForceSSL
        It -Name 'Only a single item returned' -test {
            $FunctionResult.Count | Should -Be 1
        }
        It -name 'The correct item is returned' -test {
            $FunctionResult | Should -Match ('^https*:\/\/.+{0}$' -f $SingleFileNameRegEx)
        }
        It -name 'Returned item is of System.String type' -test {
            $FunctionResult | Should -BeOfType 'System.String'
        }
    }
    Context -Name 'Multiple files - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $MultipleFilesGUID -LinksOnly -ForceSSL
        It -name 'Multiple files returned by GUID' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
    Context -Name 'SC - Multiple files - by ID' -Fixture {
        $FunctionResult = &($FunctionName) -KB $MultipleFilesKBIDSC -SearchCriteria $SearchCriteriaSC -LinksOnly -ForceSSL
        It -name 'Multiple items returned' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesSCRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesSCRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
    Context -Name 'SC - Multiple files - by GUID' -Fixture {
        $FunctionResult = &($FunctionName) -GUID $MultipleFilesGUIDSC -LinksOnly -ForceSSL
        It -name 'Multiple files returned by GUID' -test {
            $FunctionResult.Count | Should -Be 2
        }
        It -name 'The correct items are returned' -test {
            $FRContains = 0
            foreach ($Item in $FunctionResult) {
                foreach ($Item2 in $MultipleFilesNamesSCRegEx) {
                    if ($Item -match ('^https*:\/\/.+{0}$' -f $Item2)) {
                        $FRContains++
                    }
                }
            }
            $EXContains = 0
            foreach ($Item in $MultipleFilesNamesSCRegEx) {
                foreach ($Item2 in $FunctionResult) {
                    if ($Item2 -match ('^https*:\/\/.+{0}$' -f $Item)) {
                        $EXContains++
                    }
                }
            }
            $EXContains | Should -Be $FunctionResult.Count
            $FRContains | Should -Be $FunctionResult.Count
        }
        It -name 'Returned items are of System.String type' -test {
            foreach ($Item in $FunctionResult) {
                $Item | Should -BeOfType 'System.String'
            }
        }
    }
}

Describe -Name 'Unit tests' -Fixture {
    Push-Location -Path 'TestDrive:\'
    #$UnitFunctionName = 'FindTableColumnNumber'
    #. ([scriptblock]::Create((([System.Management.Automation.Language.Parser]::ParseInput((Get-Content function:$FunctionName), [ref]$null, [ref]$null)).beginBlock.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true) | Where-Object -FilterScript {$_.Name -eq $UnitFunctionName}).Extent.Text))
    #& ($UnitFunctionName) -Columns -Pattern

    #$UnitFunctionName = 'GetKBDownloadLinksByGUID'
    #. ([scriptblock]::Create((([System.Management.Automation.Language.Parser]::ParseInput((Get-Content function:$FunctionName), [ref]$null, [ref]$null)).beginBlock.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true) | Where-Object -FilterScript {$_.Name -eq $UnitFunctionName}).Extent.Text))
    #& ($UnitFunctionName) -Columns -Pattern

    Context -Name 'DownloadWUFile' -Fixture {
        $UnitFunctionName = 'DownloadWUFile'
        . ([scriptblock]::Create((([System.Management.Automation.Language.Parser]::ParseInput((Get-Content function:$FunctionName), [ref]$null, [ref]$null)).EndBlock.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true) | Where-Object -FilterScript {$_.Name -eq $UnitFunctionName}).Extent.Text))

        Context ('{0} - Single file - HTTP' -f $UnitFunctionName) -Fixture {
            $FunctionResult = & ($UnitFunctionName) -URL $HTTPSingleString -DestinationFolder '.\' -FileName $SingleFileName

            It -name 'All files are downloaded' -test {
                foreach ($Item in $SingleFileName) {
                    $Item | Should -Exist
                }
            }
            It -name 'Only a single item returned' -test {
                $FunctionResult.Count | Should -Be 1
            }
            It -name 'Returned items are of System.IO.FileInfo type' -test {
                foreach ($Item in $FunctionResult) {
                    $Item | Should -BeOfType 'System.IO.FileInfo'
                }
            }

            foreach ($Item in $SingleFileName) {
                Remove-Item -Path $Item -Force
            }
        }

        <# Unfortunately, download.microsoft.com has an incorrect certificate installed at some nodes, which renders HTTPS testspop useless.
        Context ('{0} - SingleFile - HTTPS' -f $UnitFunctionName) -Fixture {
            $FunctionResult = & ($UnitFunctionName) -URL $HTTPSSingleString -DestinationFolder '.\' -FileName $SingleFileName

            It -name 'All files are downloaded' -test {
                foreach ($Item in $SingleFileName) {
                    $Item | Should -Exist
                }
            }
            It -name 'Only a single item returned' -test {
                $FunctionResult.Count | Should -Be 1
            }
            It -name 'Returned items are of System.IO.FileInfo type' -test {
                foreach ($Item in $FunctionResult) {
                    $Item | Should -BeOfType 'System.IO.FileInfo'
                }
            }

            foreach ($Item in $SingleFileName) {
                Remove-Item -Path $Item -Force
            }
        } #>

        Context ('{0} - Multiple files - HTTP' -f $UnitFunctionName) -Fixture {
            $FunctionResult = @()
            foreach ($Item in $HTTPMultiple) {
                $Item -match '.+/(.+)$'
                $FunctionResult += & ($UnitFunctionName) -URL $Item -DestinationFolder '.\' -FileName $Matches[1]
            }

            It -name 'All files are downloaded' -test {
                foreach ($Item in $MultipleFilesNames) {
                    $Item | Should -Exist
                }
            }
            It -name 'Multiple items returned' -test {
                $FunctionResult.Count | Should -Be 2
            }
            It -name 'Returned items are of System.IO.FileInfo type' -test {
                foreach ($Item in $FunctionResult) {
                    $Item | Should -BeOfType 'System.IO.FileInfo'
                }
            }
    
            foreach ($Item in $MultipleFilesNames) {
                Remove-Item -Path $Item -Force
            }
        }
    }

    Context -Name 'RewriteURLtoHTTPS' -Fixture {
        $UnitFunctionName = 'RewriteURLtoHTTPS'
        . ([scriptblock]::Create((([System.Management.Automation.Language.Parser]::ParseInput((Get-Content function:$FunctionName), [ref]$null, [ref]$null)).EndBlock.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true) | Where-Object -FilterScript {$_.Name -eq $UnitFunctionName}).Extent.Text))

        It -name ('{0} - HTTPMultiple' -f $UnitFunctionName) -test {
            & ($UnitFunctionName) -URL $HTTPMultiple | Should -Be $HTTPSMultiple
        }
        It -name ('{0} - HTTPSMultiple' -f $UnitFunctionName) -test {
            & ($UnitFunctionName) -URL $HTTPSMultiple | Should -Be $HTTPSMultiple
        }
        It -name ('{0} - HTTPSingle' -f $UnitFunctionName) -test {
            & ($UnitFunctionName) -URL $HTTPSingle | Should -Be $HTTPSSingle
        }
        It -name ('{0} - HTTPSSingle' -f $UnitFunctionName) -test {
            & ($UnitFunctionName) -URL $HTTPSSingle | Should -Be $HTTPSSingle
        }
        It -name ('{0} - HTTPSingleString' -f $UnitFunctionName) -test {
            & ($UnitFunctionName) -URL $HTTPSingleString | Should -Be $HTTPSSingle
        }
        It -name ('{0} - HTTPSSingleString' -f $UnitFunctionName) -test {
            & ($UnitFunctionName) -URL $HTTPSSingleString | Should -Be $HTTPSSingle
        }
    }

    Context -Name 'ParseKBDownloadLinksFromText' -Fixture {
        $UnitFunctionName = 'ParseKBDownloadLinksFromText'
        . ([scriptblock]::Create((([System.Management.Automation.Language.Parser]::ParseInput((Get-Content function:$FunctionName), [ref]$null, [ref]$null)).EndBlock.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true) | Where-Object -FilterScript {$_.Name -eq $UnitFunctionName}).Extent.Text))
        
        It -name ('{0} - Multiple' -f $UnitFunctionName) -test {
            $TextMultiple = @"
    //writes the eula and file info arrays
    var downloadInformation = new Array();
downloadInformation[0] = new Object();
downloadInformation[0].updateID ='cdde339c-ebdb-4a16-add4-fb196a5053a8';
downloadInformation[0].isHotFix =false;
downloadInformation[0].enTitle ='Security Update for Windows Server 2012 R2 (KB3172729)';
downloadInformation[0].sizeLanguage ='';
downloadInformation[0].files = new Array();
downloadInformation[0].files[0] = new Object();
downloadInformation[0].files[0].url = 'http://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu';
downloadInformation[0].files[0].digest = '';
downloadInformation[0].files[0].architectures = 'AMD64';
downloadInformation[0].files[0].languages = 'all';
downloadInformation[0].files[0].longLanguages = 'all';
downloadInformation[0].files[0].fileName = 'windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu';
downloadInformation[0].files[0].defaultFileNameLength = 141;
downloadInformation[0].files[1] = new Object();
downloadInformation[0].files[1].url = 'http://download.windowsupdate.com/d/msdownload/update/software/secu/2016/07/windows8.1-kb3172729-x64_e8003822a7ef4705cbb65623b72fd3cec73fe222.msu';
downloadInformation[0].files[1].digest = '';
downloadInformation[0].files[1].architectures = 'AMD64';
downloadInformation[0].files[1].languages = 'all';
downloadInformation[0].files[1].longLanguages = 'all';
downloadInformation[0].files[1].fileName = 'windows8.1-kb3172729-x64_e8003822a7ef4705cbb65623b72fd3cec73fe222.msu';
downloadInformation[0].files[1].defaultFileNameLength = 141;
downloadInformation[0].allFilesExist = false;
var minFilePathLength =137;
var eulaInfo =  new Array();


"@
            $TestResultMultiple = @(
                'http://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu',
                'http://download.windowsupdate.com/d/msdownload/update/software/secu/2016/07/windows8.1-kb3172729-x64_e8003822a7ef4705cbb65623b72fd3cec73fe222.msu'
            )

            $Result = & ($UnitFunctionName) -Text $TextMultiple -KB '3173424'

            #Can't wait for array comparison in Pester 4.1
            foreach ($Item in $Result) {
                $TestResultMultiple -contains $Item | Should -Be $true
            }
            foreach ($Item in $TestResultMultiple) {
                $Result -contains $Item | Should -Be $true
            }
        }
        It -name ('{0} - Single' -f $UnitFunctionName) -test {
            $TextSingle = @"
    //writes the eula and file info arrays
    var downloadInformation = new Array();
downloadInformation[0] = new Object();
downloadInformation[0].updateID ='cdde339c-ebdb-4a16-add4-fb196a5053a8';
downloadInformation[0].isHotFix =false;
downloadInformation[0].enTitle ='Security Update for Windows Server 2012 R2 (KB3172729)';
downloadInformation[0].sizeLanguage ='';
downloadInformation[0].files = new Array();
downloadInformation[0].files[0] = new Object();
downloadInformation[0].files[0].url = 'http://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu';
downloadInformation[0].files[0].digest = '';
downloadInformation[0].files[0].architectures = 'AMD64';
downloadInformation[0].files[0].languages = 'all';
downloadInformation[0].files[0].longLanguages = 'all';
downloadInformation[0].files[0].fileName = 'windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu';
downloadInformation[0].files[0].defaultFileNameLength = 141;
downloadInformation[0].allFilesExist = false;
var minFilePathLength =137;
var eulaInfo =  new Array();


"@
            $TestResultSingle = 'http://download.windowsupdate.com/d/msdownload/update/software/crup/2016/06/windows8.1-kb3173424-x64_9a1c9e0082978d92abee71f2cfed5e0f4b6ce85c.msu'
            
            $Result = & ($UnitFunctionName) -Text $TextSingle -KB '3173424'
            $Result | Should -Be $TestResultSingle
        }
    }
    Pop-Location
}

Describe -Name 'Errors' -Fixture {
    Push-Location -Path 'TestDrive:\'
    Context -Name 'KB ID is incorrect' -Fixture {
        $Result = $false
        try {
            & ($FunctionName) -KB '11111111' -SearchCriteria $SearchCriteria -ErrorAction 'Stop'
            $Result = $true
        }
        catch {
            It -name 'Error category is correct' -test {
                $_.CategoryInfo.Category | Should -Be ([System.Management.Automation.ErrorCategory]::InvalidOperation)
            }
            It -name 'Exception fully qualified error id is correct' -test {
                $_.Exception.ErrorRecord.FullyQualifiedErrorId | Should -Be 'InvokeMethodOnNull'
            }
            It -name 'Inner exception fully qualified error id is correct' -test {
                $_.Exception.InnerException.ErrorRecord.FullyQualifiedErrorId | Should -Be 'InvokeMethodOnNull'
            }
        }
        It -name 'The function should fail' -test {
            $Result | Should -Be $false
        }
    }

    Context -Name 'Catalog DNS name could not be resolved' -Fixture {
        $Result = $false
        try {
            & ($FunctionName) -KB $SingleFileKBID -SearchCriteria $SearchCriteria -SearchPageTemplate 'https://www.atalog.update.microsoft.com/Search.aspx?q={0}' -ErrorAction 'Stop'
            $Result = $true
        }
        catch {
            It -name 'Error object target object URI property is correct' -test {
                $_.TargetObject.RequestUri | Should -Be ('https://www.atalog.update.microsoft.com/Search.aspx?q={0}' -f $SingleFileKBID)
            }
            It -name 'Error object target object Host property is correct' -test {
                $_.TargetObject.Host | Should -Be 'www.atalog.update.microsoft.com'
            }
            It -name 'Error object target object type is correct' -test {
                $_.TargetObject | Should -BeOfType 'System.Net.HttpWebRequest'
            }
            It -name 'Error object inner exception status is correct' -test {
                $_.Exception.InnerException.Status | Should -Be ([System.Net.WebExceptionStatus]::NameResolutionFailure)
            }
        }
        It -name 'The function should fail' -test {
            $Result | Should -Be $false
        }
    }
    Pop-Location
}

Describe -Name 'Comment-based help' -Fixture { # http://www.lazywinadmin.com/2016/05/using-pester-to-test-your-comment-based.html
    $Help = Get-Help -Name $FunctionName -Full
    $Notes = ($Help.alertSet.alert.text -split '\n')

    Context -Name ('{0} - Help' -f $FunctionName) -Fixture {
            
        It -name 'Synopsis' -test {
            $help.Synopsis | Should -Not -BeNullOrEmpty
        }
        It -name 'Description' -test {
            $help.Description | Should -Not -BeNullOrEmpty
        }
        It -name 'Notes - Author' -test {
            $Notes[0].trim() | Should -Be 'Author: Kirill Nikolaev'
        }
        It -name 'Notes - Twitter' -test {
            $Notes[1].trim() | Should -Be 'Twitter: @exchange12rocks'
        }
        It -name 'Notes - Web-site' -test {
            $Notes[2].trim() | Should -Be 'Web-site: https://exchange12rocks.org'
        }
        It -name 'Notes - GitHub' -test {
            $Notes[3].trim() | Should -Be 'GitHub: https://github.com/exchange12rocks'
        }

        # Get the parameters declared in the Comment Based Help
        $RiskMitigationParameters = 'Whatif', 'Confirm'
        $HelpParameters = $help.parameters.parameter | Where-Object name -NotIn $RiskMitigationParameters

        # Parse the function using AST
        $AST = [System.Management.Automation.Language.Parser]::ParseInput((Get-Content function:$FunctionName), [ref]$null, [ref]$null)

        # Get the parameters declared in the AST PARAM() Block
        $ASTParameters = $AST.ParamBlock.Parameters.Name.variablepath.userpath

        It -name 'Parameter - Compare Count Help/AST' -test {
            $HelpParameters.name.count | Should -Be $ASTParameters.count
        }
            
        # Parameter Description
        If (-not [String]::IsNullOrEmpty($ASTParameters)) {
            # IF ASTParameters are found
            $HelpParameters | ForEach-Object {
                It -name ('Parameter {0} - Should contains description' -f $_.Name) -test {
                    $_.description | Should -Not -BeNullOrEmpty
                }
            }
        }
            
        # Examples
        It -name 'Example - Count should be greater than 0' -test {
            $Help.examples.example.count | Should -BeGreaterThan 0
        }
        
        # Every parameter set should be covered by at least one example, but since I do not see a better way to test it, let's just count the number of examples and compare it to the number of parameter sets.
        It -name 'Examples - At least one example per ParameterSet' -test {
            $Help.examples.example.count | Should -BeGreaterThan ($Help.syntax.syntaxItem.Count-1)
        }
        
        # Examples - Code ("code" is the first line of an example)
        foreach ($Example in $Help.examples.example) {
            It -name ('Example - Code on {0}' -f $Example.Title) {
                $Example.code | Should -Not -Be '' # There is no reason to leave the first row of an example blank
            }
        }

        # Examples - Remarks (small description that comes with the example)
        foreach ($Example in $Help.examples.example) {
            It -name ('Example - Remarks on {0}' -f $Example.Title) {
                if ($Example.remarks -is 'System.Array') {
                        $Example.remarks[0] | Should -Not -Be '' # Strangely, remarks section is usually an array of 5 elements where only the first one contains text
                }
                else {
                    $Example.remarks | Should -Not -BeNullOrEmpty
                }
            }
        }
    }
}