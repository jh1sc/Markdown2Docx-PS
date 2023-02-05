
if ([bool](winget install pandoc) -ne $true) { winget install pandoc }
Function FOSD {
    [CmdletBinding()]
    Param(
        [string]$Description = "Select a folder",
        [string]$RootFolder = "Desktop"
    )
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.Description = $Description
    $FolderBrowserDialog.RootFolder = $RootFolder
    $FolderBrowserDialog.ShowDialog() | Out-Null
    return $FolderBrowserDialog.SelectedPath
}
Function FISD {
    Param(
        [string]$Title = "Select a file",
        [string[]]$Filter = @("All Files (*.*)|*.*")
    )
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $Title
    $OpenFileDialog.Filter = ($Filter -join "|")
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
}

$MarkDown = FISD -Title "Choose a .md file" -Filter @("Markdown Files (*.md)|*.md")
$Directory = FOSD -Description "Select a directory to create the new docx in"
$Directory = $Directory+"\MarkDown.docx"
.\pandoc -s $MarkDown -o $Directory --reference-doc .\ref.docx
Start-Process $Directory 


