<#
.SYNOPSIS
    PowerCountPST
.DESCRIPTION
    This PowerShell Script lets you count the number of elements found in an Outlook PST file. There are no limitations in terms of size and it's blazingly fast.
    Unlike other solutions, this script actually searches recursively, so no matter how your files PST files are structured, it will always yield back the accurate amount.
.PARAMETER pst
    The full qualified path to the PST file.
.EXAMPLE
    .\PowerCountPST.ps1 -pst "C:\Temp\outlook.pst"
    This example show how to pass a pst file located in the temp folder to the script.
.LINK https://github.com/lxndrblz/PowerCountPST/
.NOTES
    Author: Alexander Bilz
    Date:   April 18, 2020    
#>

[CmdletBinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = "Low"
)]

Param
(
    [Parameter(Mandatory = $true, Position = 0, ParameterSetName = "pst", 
        ValueFromPipeline = $true, 
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Path to the PST file")]
    [ValidateNotNullOrEmpty()]
    [string]
    $pst
)
process {

    # Errors
    $ERR_FILENOTFOUND = 1000
    $ERR_ROOTFOLDER = 1001
    $ERR_LOCKEDFILE = 1002
    $ERR_CANTACCESSPST = 1003

    function AnalyzeFolder ($folder) {
    
        $count = $folder.items.count
        Write-Host ('Folder: {0} contains {1} items' -f $folder.FolderPath, $count)
        Foreach ($subfolder in $folder.Folders) {
            $subfolderitems = AnalyzeFolder $subfolder
            $count = $count + $subfolderitems
        }

        return $count
    }

    function CountElements ($strPSTPath) {

        $counter = 0

        #Check if Outlook is installed 
        Get-ItemProperty HKLM:\SOFTWARE\Classes\Outlook.Application | Select-Object PSPath -OutVariable outlook 
 
        #Create Outlook COM Object 
        $objOutlook = New-Object -com Outlook.Application 
        $objNameSpace = $objOutlook.GetNamespace("MAPI") 
 
        #Try to load the PST into Outlook 
        try { 
            $objNameSpace.AddStore($strPSTPath) 
        } 
        catch { 
            Write-Error "Could not load pst - usually this is because the file is locked by another process."
            Exit $ERR_LOCKEDFILE
        } 
 
        #Try to load the Outlook Folders 
        try { 
            $PST = $objnamespace.stores | ? { $_.FilePath -eq $strPSTPath } 
        } 
        catch { 
            Write-Error "You have another PST added to outlook that cannot be accessed or found, please remove then re-run this script."
            Exit $ERR_CANTACCESSPST
        }

        #Try accessing the PST root
        try { 
            #Browse to PST Root 
            $root = $PST.GetRootFolder() 
   
            #Count Items in subfolders
            $counter = AnalyzeFolder $root

            # Output total number of found elements
            Write-Host $strPSTPath
            Write-Host ('Total Items: {0}' -f $counter)

            # Unmount PST
            $objNameSpace.RemoveStore($root) 
            $objOutlook.Quit() | out-null 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objOutlook)
            Remove-Variable outlook | out-null
            return $counter
        }
        catch { 
            Write-Error "Could not access root folder"
            Exit $ERR_ROOTFOLDER
        }
    }
    If (Test-Path $pst) {
        CountElements($pst)
    }
    Else {
        Write-Error "Please provide a valid path to a pst file."
        Exit $ERR_FILENOTFOUND
    }

}

