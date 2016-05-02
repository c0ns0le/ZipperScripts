$shell=new-object -com shell.application
$CurrentLocation=get-location
$CurrentPath=$CurrentLocation.path
$Location=$shell.namespace($CurrentPath)
$ZipFiles = get-childitem *.zip
$ZipFiles.count | out-default
foreach ($ZipFile in $ZipFiles)
{
$ZipFile.fullname | out-default
$ZipFolder = $shell.namespace($ZipFile.fullname)
$Location.Copyhere($ZipFolder.items())
}