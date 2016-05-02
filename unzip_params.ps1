$source = Read-Host -Prompt 'Insert source path'
$destination = Read-Host -Prompt 'Insert destination path' 

$shell=new-object -com shell.application
$CurrentLocation=$source
$CurrentPath=$destination
$Location=$shell.namespace($CurrentPath)
$ZipFiles = get-childitem *.zip
$ZipFiles.count | out-default
foreach ($ZipFile in $ZipFiles)
{
$ZipFile.fullname | out-default
$ZipFolder = $shell.namespace($ZipFile.fullname)
$Location.Copyhere($ZipFolder.items())
}