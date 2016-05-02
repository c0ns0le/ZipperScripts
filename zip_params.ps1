# 
# Powershell function (Works in 1.0 and 2.0) 
# to create a .ZIP file using the native "Send to Compressed Folder" 
# feature in Windows explorer.   It then copies files into the "Folder" thus allowing you 
# to create a Zip archive natively in Powershell 
# 
# Original code from a post by Steve Fulton on 
# http://social.msdn.microsoft.com/Forums/en-US/windowsgeneraldevelopmentissues/thread/d3e347cc-f4dc-44a6-8f84-977f958d89c6/ 
# Using VB.Net 
#  
# Orignal Notes and Code from Steve Fulton Follows.... 
#  
#  1. Copy it EXACTLY as shown, make sure it works, then modify it. For example,  
#  "ToString" shouldn't be required, but is.. 
#  
#  2. This has only been tested on Windows Vista and Windows Server 2003 
#  It should work on any version of Windows that pretends zip files are compressed folders 
# 
#  3. This works by creating an empty zip file. Then it copies the file you want to compress into the zip file. Windows will #see that you are copying into a zip file not a folder, and compress it on the way in. 
# 
#  ------------------------------------------------------------- 
# 
#  Private Sub zipFile(ByVal filename As String, ByVal zipfilename As String) 
# 
#        Dim strZIPHeader As String 
# 
#        strZIPHeader = [char]80 + [char]75 + [char]5 + [char]6 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0) 
#        Dim fso = CreateObject("Scripting.FileSystemObject") 
#        Dim tf = fso.CreateTextFile(zipfilename) 
#        tf.Write(strZIPHeader) 
#        tf.Close() 
# 
#        With CreateObject("Shell.Application") 
#            .NameSpace(zipfilename.ToString).CopyHere(filename.ToString) 
#        End With 
#        MessageBox.Show("All done!") 
#    End Sub 
# 
# ------------------------------------------------------------- 
# 
# Sean (Energized Tech) Notes. 
#  
# I found this immensely cool to play with since (If you look at the Structure) not a 
# lot changes between VB.Net and Powershell suggesting it would not be hard to 
# translate many functions or applications from VB.Net to Powerhshell 
# 
 
function global:SEND-ZIP ($zipfilename, $filename) { 
 
<# 
 
.SYNOPSIS  
Function to send Files / Folders to a ZIP file using the native 
feature in Windows Vista / 7 / Windows XP 
 
.DESCRIPTION  
Function to send Files / Folders to a ZIP file using the native 
feature in Windows Vista / 7 / Windows XP.  Requires 
Two parameters to be sent.  The Zip file name (with .ZIP) and 
A File / Folder 
 
.EXAMPLE  
Send a Folder called C:\FolderA to a file in the current 
folder called MYZIPFILE.ZIP 
 
SEND-ZIP C:\MYZIPFILE.ZIP C:\Foldera 
 
You must ALWAYS Specify an EXPLICIT path to the ZIP file 
 
.EXAMPLE  
Send a File called FILE.TXT in C:\FOLDER to a  
ZIP file called TEMPZIPFILE.ZIP in the C:\TEMP Folder 
 
SEND-ZIP C:\TEMP\NEWZIPFILE.ZIP C:\FOLDER\FILE.TXT 
 
You must ALWAYS Specify an EXPLICIT path to the ZIP file 
 
.EXAMPLE  
This will FAIL - Consider it a flaw in design :) 
 
SEND-ZIP NEWZIP.ZIP C:\Foldera 
 
.NOTES  
This was originally a VB.Net Function written by Steve Fulton 
from a post on MSDN.COM 
http://social.msdn.microsoft.com/Forums/en-US/windowsgeneraldevelopmentissues/thread/d3e347cc-f4dc-44a6-8f84-977f958d89c6/ 
 
Converted to Powershell by Sean Kearney @energizedtech (www.powershell.ca) 
 
#> 
# The $zipHeader variable contains all the data that needs to sit on the top of the  
# Binary file for a standard .ZIP file 
$zipHeader=[char]80 + [char]75 + [char]5 + [char]6 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 
 
# 
# Check to see if the Zip file exists, if not create a blank one 
# 
If ( (TEST-PATH $zipfilename) -eq $FALSE ) { Add-Content $zipfilename -value $zipHeader } 
 
# Create an instance to Windows Explorer's Shell comObject 
# 
$ExplorerShell=NEW-OBJECT -comobject 'Shell.Application' 
 
# 
# Send whatever file / Folder is specified in $filename to the Zipped folder $zipfilename 
# 
$SendToZip=$ExplorerShell.Namespace($zipfilename.tostring()).CopyHere($filename.ToString()) 
 
# 
# Hey the cool part is if you send a folder, it automatically recurses and gives you the 
# Progress bar as it is zipping. 
# 
# Sean 
# The Energized Tech 
# www.powershell.ca 
 
} 