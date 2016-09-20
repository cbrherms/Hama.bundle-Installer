	Plex 	= InputBox("The path where local application data is stored (from Settings)") 
	''MsgBox Plex
	'Plex = Plex & "Plex Media Server\"
	'MsgBox Plex
	
	DestF 	= "C:\TempPlex"
	DestF2 	= "C:\TempPlex2"
	File1 	= "Hamma.Scanner"
	myURL1 	= "https://github.com/ZeroQI/Absolute-Series-Scanner/archive/master.zip"
	File2 	= "Hamma.Plugin"
	myURL2 	= "https://github.com/ZeroQI/Hama.bundle/archive/master.zip"
	File3 	= "Hamma.PluginSupport"
	myURL3 	= "https://github.com/ZeroQI/Hama.bundle/releases/download/v1.0/Plug-in.Support.zip"
	
	PlexTest = "C:\TempPlex\Plex Media Server"
	
	CreateFolder DestF
	'CreateFolder DestF2
	
	CreateFolderIf Plex
	CreateFolderIf Plex & "\Scanners\"
	CreateFolderIf Plex & "\Scanners\Series\"
	CreateFolderIf Plex & "\Plug-ins\"
	CreateFolderIf Plex & "\Plug-in Support\"

	Procesing myURL1,File1,DestF & "\"
	Procesing myURL2,File2,DestF & "\"
	Procesing myURL3,File3,DestF & "\"
	
	'Copy1
	CFrom1 	=	DestF & "\" & File1 & "\Absolute-Series-Scanner-master\Scanners\Series\Absolute Series Scanner.py"
	CTo1	=	Plex & "\Scanners\Series\" 'Plex & "\Scanners\Series\"
	'MsgBox CFrom1
	'MsgBox CTo1
	CopyFile CFrom1, CTo1
	
	'Copy2
	CFrom2 	=	DestF & "\" & File2 & "\Hama.bundle-master\Contents"
	CTo2	=	Plex & "\Plug-ins\Hama.bundle" 'Plex & "\Plug-ins\Hama.bundle"
	'MsgBox CFrom2
	'MsgBox CTo2
	CopyFolder CFrom2, CTo2
	
	'Copy3
	CFrom3 	=	DestF & "\" & File3 & "\Plug-in Support"
	CTo3	=	Plex & "\Plug-in Support" 'Plex
	'MsgBox CFrom3
	'MsgBox CTo3
	CopyFolder CFrom3, CTo3
	
	'MsgBox "Deleting Temo Folder"
	DeleteFolder DestF
	'MsgBox "Deleted temp folder"
'************************************************************************************************************
	Sub Procesing(myURL,ImageFile,DestFolder)
		Zip = ".zip"
		myPath = DestFolder & ImageFile  & Zip
		HTTPDownload myURL, myPath	
		UnZip myPath, DestFolder & ImageFile
	End Sub
	
	Sub CopyFolder(Source, Dest)
		Const OverWriteFiles = True
		Set objFSO = CreateObject("Scripting.FileSystemObject")
			If Not objFSO.FolderExists(Dest) Then
				objFSO.CreateFolder(Dest) 'CreateFolder Dest
			End If
		objFSO.CopyFolder Source , Dest , OverWriteFiles
	End Sub
	
	Sub CopyFile(Source, Dest)
		Const OverWriteFiles = True
		Set objFSO = CreateObject("Scripting.FileSystemObject")
			If Not objFSO.FolderExists(Dest) Then
				objFSO.CreateFolder(Dest)
			End If
		objFSO.CopyFile Source , Dest , OverWriteFiles	
	End Sub
	
	Sub UnZip(ZipFile,ExtractTo)
		'If the extraction location does not exist create it.
		Set fso = CreateObject("Scripting.FileSystemObject")
		If NOT fso.FolderExists(ExtractTo) Then
		   fso.CreateFolder(ExtractTo)
		End If

		'Extract the contants of the zip file.
		set objShell = CreateObject("Shell.Application")
		set FilesInZip=objShell.NameSpace(ZipFile).items
		objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)
		Set fso = Nothing
		Set objShell = Nothing
	End Sub		
	
	Sub CreateFolder(FolderToBeCreated)
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.CreateFolder(FolderToBeCreated)
		CreateFolderDemo = f.Path
	End Sub
	
	Sub CreateFolderIf(FolderToBeCreated)
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Not fso.FolderExists(FolderToBeCreated) Then
			Set f = fso.CreateFolder(FolderToBeCreated)
		End If
		'CreateFolderDemo = f.Path
	End Sub
	
	Sub DeleteFolder(filespec)
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		'MsgBox filespec
		fso.DeleteFolder(filespec)
	End Sub
	
	Sub HTTPDownload( myURL, myPath )
	' This Sub downloads the FILE specified in myURL to the path specified in myPath.
	'
	' myURL must always end with a file name
	' myPath may be a directory or a file name; in either case the directory must exist
	'
	' Written by Rob van der Woude
	' http://www.robvanderwoude.com
	'
	' Based on a script found on the Thai Visa forum
	' http://www.thaivisa.com/forum/index.php?showtopic=21832

		' Standard housekeeping
		Dim i, objFile, objFSO, objHTTP, strFile, strMsg
		Const ForReading = 1, ForWriting = 2, ForAppending = 8

		' Create a File System Object
		Set objFSO = CreateObject( "Scripting.FileSystemObject" )

		' Check if the specified target file or folder exists,
		' and build the fully qualified path of the target file
		If objFSO.FolderExists( myPath ) Then
			strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
		ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
			strFile = myPath
		Else
			WScript.Echo "ERROR: Target folder not found."
			Exit Sub
		End If

		' Create or open the target file
		Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )

		' Create an HTTP object
		Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )

		' Download the specified URL
		objHTTP.Open "GET", myURL, False
		objHTTP.Send

		' Write the downloaded byte stream to the target file
		For i = 1 To LenB( objHTTP.ResponseBody )
			objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
		Next

		' Close the target file
		objFile.Close( )
	End Sub
