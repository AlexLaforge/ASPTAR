## ASPTAR ##

UNIX Tarball creator in VBScript

What is ASPTAR?
ASPTAR is an ASP implementation of the UNIX tarball file format. TAR stands for Tape ARchive. This is a simple means of grouping multiple files together in a single file for easier transportation. It's kind of like a zip file, but it isn't compressed. It enables your application to distribute multiple files with a single download. 

What supports the TAR file format?
A number of commercially available applications support the TAR file format, including the very popular WinZip, which allows you to extract data from a TAR archive file. You will also find that the UNIX operating system supports this file type out of the box. 

How does it work?
ASPTAR works by reading data from both disk files and also direct from memory, then it creates the necessary file format expected by TAR extraction programs. This is then sent to the browser which prompts the user to save the file locally prior to extraction. 

Requirements

IIS 5.0
VBScript 5.6
MDAC 2.6

These are all important as the internals of ASPTAR have changed to utilise ADO Streams
Also note that this class will still NOT work under ChiliSoft's ASP package.

Installation instructions
Copy the asptar.asp file to your web server. You can include asptar.asp at the top of your pages using SSI. See the example in this document on how to use the class.

Usage
It is important to remember the rules of client/server applications. If you are to use this class within your web based application, it should be used on a page which only outputs the TAR file. Do not attempt to make a page which outputs both HTML and the TAR data, it simply won't work! With that explanation out of the way, here's a quick example.

```
<!--#include file="asptar.asp"-->
<%
Dim objTar
Set objTar = New Tarball
objTar.AddMemoryFile "mum.txt","Hello mum!"
objTar.AddMemoryFile "dad.txt","Hello dad!"
objTar.WriteTar
%>
```

Save this to an ASP page and place it in the same folder as the asptar.asp program.
Running the page will show you a "Save As" dialog. Save it somewhere on your computer and then open the .tar file with an application like WinZip. Lets explain the above example: The first line is a Server Side Include to make sure that the ASPTAR class is present for ASP to use. The fourth line actually creates an instance of the tarball in memory, so you can do things to it. The fifth and sixth lines add a file to the tarball archive from two strings "Hello mum!" and "Hello dad!", each called mum.txt and dad.txt respectively. The seventh line actually sends the tarball file to the browser for someone to download. Normally, the page that this code would be saved in would be linked from another page with an anchor reference. 

Methods
AddFile - AddFile filename - Adds a disk file location to the archive
RemoveFile - RemoveFile filename - Removes a file location from the archive
AddMemoryFile - AddMemoryFile filename,contents - Adds a file to the archive from a string
RemoveMemoryFile - RemoveMemoryFile filename - Removes a file from the archive
WriteTar - WriteTar - Sends the tarball to the browser for downloading

Properties
TarFilename - String - Name of the saved tarball (default=new.tar)
UserID - Long - UNIX user ID number (not used in Windows)
GroupID - Long - UNIX group ID number (not used in Windows)
UserName - String - UNIX user name (not used in Windows)
GroupName - String - UNIX group name (not used in Windows)
Permissions - Long - UNIX permissions settings (default=438, or 666 in octal) (not used in Windows)
BlockSize - Long - TAR block byte size (default=512)
IgnorePaths - True/False - When writing the TAR file, don't include the source path names for each disk file
BasePath - String - Set a path to include in with the archive which can be used to set an extraction directory
(IgnorePaths needs to be true if this is to be used)

Current limitations

ChiliSoft's ASP package still doesn't support VBScript classes, so this wont work with that MDAC 2.6 is required so that ADO Streams are present and correct

Technical information

All header information for TAR files is stored in ASCII in octal ADO Streams are utilised to increase the speed of execution

ASP Client-side debugging
If you have this option turned on in IIS for the web site which ASPTAR is  running on, the integration between IE and IIS will cause IE to corrupt the archive.  Ensure that client-side debugging is disabled in the IIS MMC before testing ASPTAR.  You can find this option in Default Web Site->Properties->Home->  directory->Configuration->App debugging->Enable ASP client-side script  debugging. 

Setting the LCID
Due to the way in which ASPTAR internally renders the tarball, it has become necessary to make sure that certain conditions are met if you are developing for a different Locale ID other than English. It is advised that the LCID of the Session object be set to the value 1033 (US English) for just the tarball generation page. This ensures that ADO will interpret the x-ansi stream correctly  without any character conversions. 

ChiliSoft ASP package
Please note that the ChiliSoft ASP package does not support VBScript classes, and will not run this script correctly
