Dim t2pobj, cmd, srvOpts

' Make sure to register the dll first, if necessary (your path might be different)
' C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase c:pdfrw_20.dll
' cscript.exe (or wscript.exe) rwtest.vbs
set t2pobj = CreateObject("FyTek.Text2PDF")

' Running as a server is optional. You may just call the program directly if you prefer.
' Running as a server provides control over resources and single copy of the program exists
' in memory.  The server may be on any box that is accsible from the one making the request.
' 
' In this sample we are starting a server to show how it is done and then shutting it down 
' with this script finishes.  In a production environment you would likely have these in 
' separate processes - that is, have one process that only starts or stops the server.  Other
' processes access the server to build PDFs and don't shut it down when they are done.  It 
' stays running for other processes to come along and access it.

' Set the location of the executable - only needed if we are starting the server or not using 
' a server.  When a server is used, only the host and port are needed.
' t2pobj.setExe(".\text2pdf64")
 
' Needed to test server mode - in produciton, use setKeyName and setKeyCode with your values.  
' Or use licInfo method to set your license info.
' Assumes we're using the demo, in a real situation, include setKeyCode also.
' And if using an already running server, no key name is necessary.
' t2pobj.setKeyName("demo") 

' Start the server.  Normally you would do this elsewhere, in another program
' or from a shell script, etc.  You might have several servers running for load balancing.
srvOpts = t2pobj.startServer(,,,"c:\temp\t2ptestlog.txt") 

' The setServerFile method is used to pass in a file name of IP addresses and ports that
' you have one or more servers running on.  This is static for the DLL so all users accessing
' the DLL will have the same file.  The DLL will then pass requests to different servers
' depending on load.
' For example, in serverfile.dat you might have:
' # here are the servers
' 192.168.1.124 7080
' 192.168.1.125 7080
' localhost 7080
' t2pobj.setServerFile("servers.dat") 

' In a real situation where server is already running and there is no server
' setting file (setServerFile call that provides a file of one or more Report Writer servers)
' then you might need to set the values here so the program knows the box and port to use.
' For example, when the server is running off of a Linux server on box 192.168.1.124:
' t2pobj.setServer "192.168.1.124",7080
' If the Report Writer server is on a different box (not localhost for example) then you might
' want to set this option that will auto send any files from this box to the server so it has
' all the needed files to build the PDF.  Or, use the sendFileTCP method to provide the files
' one at a time.  If more than one or you have included images, fonts, then may be easier to juse
' use setAutoSendFile.
t2pobj.setAutoSendFiles()

WScript.Echo("Building PDF1 - assumes there is a file c:\temp\sample.txt")
' Now that server is running, send some samples
cmd = t2pobj.setInFile("c:\temp\sample1.txt") 

' Create the PDF and get back the byte array - or pass false to not wait for the PDF to be returned
Set d = t2pobj.buildPDF (true)
' d.getByptes() will contain the newly build PDF - we'll save to a file but you might want to 
' display on a web page or save in a database
WScript.Echo("Msg=" & d.Msg)
WScript.Echo("Pages=" & d.Pages)
WScript.Echo("bytes=" & lenB(d.Bytes))
if (StrComp(d.Msg,"OK") = 0) Then
    SaveBinaryDataTextStream "c:\temp\myfile.pdf", d.Bytes
End If

' clear out the commands for another run        
cmd = t2pobj.resetOpts()
' let's check the server status (if multiple servers, pass true to get info on all of them)
cmd = t2pobj.serverStatus(true)
WScript.Echo(cmd)

' Now lets send some text through the method instead
' Also, we'll save the PDF locally by passing the output name to the buildPDF method
cmd = t2pobj.setPageNum("Page: %p","times",10,7,.25)
cmd = t2pobj.setPageTxtFont("helvetica",20.5)
cmd = t2pobj.addText("Here is some <b>text</b>." & vbLf)
cmd = t2pobj.addText("And a second line of <u>text</u>." & vbLf)
cmd = t2pobj.addText("Continued on next page..." & vbFormFeed)
cmd = t2pobj.addText("This is on page 2." & vbLf)
cmd = t2pobj.setPre("html")
Set d = t2pobj.buildPDF (true,"c:\temp\myfile2.pdf")
WScript.Echo("Pages=" & d.Pages)

' shut down the server - typically you would leave it running for the next process to use
srvOpts = t2pobj.stopServer() 

' Helper functions for saving the file to disk
Function RSBinaryToString(xBinary)
  Dim Binary
  'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
  If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary
  
  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)
  
  If LBinary>0 Then
    RS.Fields.Append "mBinary", adLongVarChar, LBinary
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk Binary 
    RS.Update
    RSBinaryToString = RS("mBinary")
  Else
    RSBinaryToString = ""
  End If
End Function

Function SaveBinaryDataTextStream(FileName, ByteArray)
  ' This function is to write a PDF to disk but
  ' you could instead send to a web page
  ' or store in a database, etc.

  'Create FileSystemObject object
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  
    If FS.FileExists(FileName) Then
      FS.DeleteFile FileName
    End If

  'Create text stream object
  Dim TextStream
  Set TextStream = FS.CreateTextFile(FileName,ForWriting)
  
  'Convert binary data To text And write them To the file
  TextStream.Write (RSBinaryToString (ByteArray))
End Function
