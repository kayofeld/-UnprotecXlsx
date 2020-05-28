$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = "./" 
    Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
	Title = "File to Unprotect"
}
$null = $FileBrowser.ShowDialog()

if ($FileBrowser.FileName -eq "") {
 	 exit
}
$FullName = $FileBrowser.FileName
$FileName = Split-Path -Path "$($FullName)" -Leaf -Resolve
$FileLocation = Split-Path -Path "$($FullName)"
$Location = (Get-Location).Path

$null = 7z.exe x "$($FullName)" -o"$($env:TEMP)/passRmvTmpFolder"

cd "$($env:TEMP)/passRmvTmpFolder/xl"
Get-ChildItem .\*.xml -Recurse | % {
    [Xml]$xml = Get-Content $_.FullName
    $xml | Select-Xml -XPath '//*[local-name() = ''sheetProtection'']' | ForEach-Object{$null = $_.Node.ParentNode.RemoveChild($_.Node)}
	$xml | Select-Xml -XPath '//*[local-name() = ''workbookProtection'']' | ForEach-Object{$null = $_.Node.ParentNode.RemoveChild($_.Node)}
    $xml.OuterXml | Out-File $_.FullName -encoding "UTF8"
} 
cd ../../

$null = 7z.exe a -tzip "$($FileLocation)/Unprotected-$($FileName.replace('.\', ''))" '.\passRmvTmpFolder\*'
rmdir -r passRmvTmpFolder
cd $Location