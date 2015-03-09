<%
' UNIX Tarball creator
' ====================
' Author: Chris Read
' Version: 1.0.1
' ====================
' Homepage: http://users.bigpond.net.au/mrjolly/
'
' This class provides the ability to archive multiple files together into a single
' distributable file called a tarball (The TAR actually stands for Tape ARchive).
' These are common UNIX files which contain uncompressed data.
'
' So what is this useful for? Well, it allows you to effectively combine multiple
' files into a single file for downloading. The TAR files are readable and extractable
' by a wide variety of tools, including the very widely distributed WinZip.
'
' This script can include two types of data in each archive, file data read from a disk,
' and also things direct from memory, like from a string. The archives support files in 
' a binary structure, so you can store executable files if you need to, or just store
' text.
'
' This class was developed to assist me with a few projects and has grown with every
' implementation. Currently I use this class to tarball XML data for archival purposes
' which allows me to grab 100's of dynamically created XML files in a single download.
'
' There are a small number of properties and methods, which are outlined in the
' accompanying documentation.
'
Class Tarball
	Public TarFilename			' Resultant tarball filename
	
	Public UserID				' UNIX user ID
	Public UserName				' UNIX user name
	Public GroupID				' UNIX group ID
	Public GroupName			' UNIX group name
	
	Public Permissions			' UNIX permissions
	
	Public BlockSize			' Block byte size for the tarball (default=512)

	Public IgnorePaths			' Ignore any supplied paths for the tarball output
	Public BasePath				' Insert a base path with each file
	
	' Storage for file information
	Private objFiles
	Private objMemoryFiles
	
	' File list management subs, very basic stuff
	Public Sub AddFile(sFilename)
		objFiles.Add sFilename,sFilename
	End Sub
	
	Public Sub RemoveFile(sFilename)
		objFiles.Remove sFilename
	End Sub
	
	Public Sub AddMemoryFile(sFilename,sContents)
		objMemoryFiles.Add sFilename,sContents
	End Sub
	
	Public Sub RemoveMemoryFile(sFilename)
		objMemoryFiles.Remove sFilename
	End Sub

	' Send the tarball to the browser
	Public Sub WriteTar()
		Dim objStream, objInStream, lTemp, aFiles

		Set objStream = Server.CreateObject("ADODB.Stream") ' The main stream
		Set objInStream = Server.CreateObject("ADODB.Stream") ' The input stream for data
		
		objStream.Type = 2
		objStream.Charset = "x-ansi" ' Good old extended ASCII
		objStream.Open

		objInStream.Type = 2
		objInStream.Charset = "x-ansi"

		' Go through all files stored on disk first
		aFiles = objFiles.Items
		
		For lTemp = 0 to UBound(aFiles)
			objInStream.Open
			objInStream.LoadFromFile aFiles(lTemp)
			objInStream.Position = 0
			ExportFile aFiles(lTemp),objStream,objInStream
			objInStream.Close
		Next

		' Now add stuff from memory
		aFiles = objMemoryFiles.Keys
		
		For lTemp = 0 to UBound(aFiles)
			objInStream.Open
			objInStream.WriteText objMemoryFiles.Item(aFiles(lTemp))
			objInStream.Position = 0
			ExportFile aFiles(lTemp),objStream,objInStream
			objInStream.Close
		Next

		objStream.WriteText String(BlockSize,Chr(0))

		' Rewind the stream
		' Remember to change the type back to binary, otherwise the write will truncate
		' past the first zero byte character.
		objStream.Position = 0
		objStream.Type = 1
		' Set all the browser stuff
		Response.AddHeader "Content-Disposition","filename=" & TarFilename
		Response.ContentType = "application/x-tar"
		Response.BinaryWrite objStream.Read
		
		' Close it and go home
		objStream.Close
		Set objStream = Nothing
		Set objInStream = Nothing
	End Sub
	
	' Build a header for each file and send the file contents
	Private Sub ExportFile(sFilename,objOutStream,objInStream)
		Dim lStart, lSum, lTemp
		
		lStart = objOutStream.Position ' Record where we are up to
		
		If IgnorePaths Then
			' We ignore any paths prefixed to our filenames
			lTemp = InStrRev(sFilename,"\")
			if lTemp <> 0 then
				sFilename = Right(sFilename,Len(sFilename) - lTemp)
			end if
			sFilename = BasePath & sFilename
		End If
		
		' Build the header, everything is ASCII in octal except for the data
		objOutStream.WriteText Left(sFilename & String(100,Chr(0)),100)
		objOutStream.WriteText "100" & Right("000" & Oct(Permissions),3) & " " & Chr(0) 'File mode
		objOutStream.WriteText Right(String(6," ") & CStr(UserID),6) & " " & Chr(0) 'uid
		objOutStream.WriteText Right(String(6," ") & CStr(GroupID),6) & " " & Chr(0) 'gid
		objOutStream.WriteText Right(String(11,"0") & Oct(objInStream.Size),11) & Chr(0) 'size
		objOutStream.WriteText Right(String(11,"0") & Oct(dateDiff("s","1/1/1970 10:00",now())),11) & Chr(0) 'mtime (Number of seconds since 10am on the 1st January 1970 (10am correct?)
		objOutStream.WriteText "        0" & String(100,Chr(0)) 'chksum, type flag and link name, write out all blanks so that the actual checksum will get calculated correctly
		objOutStream.WriteText "ustar  "  & Chr(0) 'magic and version
		objOutStream.WriteText Left(UserName & String(32,Chr(0)),32) 'uname
		objOutStream.WriteText Left(GroupName & String(32,Chr(0)),32) 'gname
		objOutStream.WriteText "         40 " & String(4,Chr(0)) 'devmajor, devminor
		objOutStream.WriteText String(167,Chr(0)) 'prefix and leader
		objInStream.CopyTo objOutStream ' Send the data to the stream
		
		if (objInStream.Size Mod BlockSize) > 0 then
			objOutStream.WriteText String(BlockSize - (objInStream.Size Mod BlockSize),Chr(0)) 'Padding to the nearest block byte boundary
		end if
		
		' Calculate the checksum for the header
		lSum = 0		
		objOutStream.Position = lStart
		
		For lTemp = 1 To BlockSize
			lSum = lSum + (Asc(objOutStream.ReadText(1)) And &HFF&)
		Next
		
		' Insert it
		objOutStream.Position = lStart + 148
		objOutStream.WriteText Right(String(7,"0") & Oct(lSum),7) & Chr(0)
		
		' Move to the end of the stream
		objOutStream.Position = objOutStream.Size
	End Sub
	
	' Start everything off
	Private Sub Class_Initialize()
		Set objFiles = Server.CreateObject("Scripting.Dictionary")
		Set objMemoryFiles = Server.CreateObject("Scripting.Dictionary")
		
		BlockSize = 512
		Permissions = 438 ' UNIX 666
		
		UserID = 0
		UserName = "root"
		GroupID = 0
		GroupName = "root"
		
		IgnorePaths = False
		BasePath = ""
		
		TarFilename = "new.tar"
	End Sub
	
	Private Sub Class_Terminate()
		Set objMemoryFiles = Nothing
		Set objFiles = Nothing
	End Sub
End Class
%>