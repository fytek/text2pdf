# Text2PDF
FyTek Text2PDF DLL

This program is supplied as a compiled .NET DLL assembly with the executable download of Text2PDF.  You may download this source and compile yourself (instructions are at the top of the code).

The purpose is to allow a DLL method interface to the executable program text2pdf.exe or text2pdf64.exe.  This DLL, which is compiled as both a 32-bit and 64-bit version, may be used in Visual Basic program, ASP, C#, etc. to send the information for building and retreiving PDFs from Text2PDF.  This DLL also replaces the old DLL version that was limited to running Text2PDF on the same box and had occassional memory issues.

There are commands to start a Text2PDF server which loads the executable version of Text2PDF in memory and listens on the port you specify for commands.  This also allows you to have Text2PDF running on a different server for load balancing by not building the PDF on the same box as the requestor.  You may also run several instances of Text2PDF server all on different boxes if you wish and the DLL will cycle between them to handle requests.
