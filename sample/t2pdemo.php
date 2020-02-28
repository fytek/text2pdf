<?php
$pdfobj = new COM('FyTek.Text2PDF');

# Normally this is done in a different program or shell script.
# Also, you might want to start a server on a different box altogether.
$pdfobj->startServer();

# This is just one way to send commands.
# You may create a file on disk as well and the file instead.

$pdfobj->setPageNum("Page: %p","times",10,7,.25);
$pdfobj->setPageTxtFont("helvetica",20.5);
$pdfobj->addText("Here is some <b>text</b>.\n");
$pdfobj->addText("And a second line of <u>text</u>.\n");
$pdfobj->addText("Continued on next page...\f");
$pdfobj->addText("This is on page 2 - a plain old form feed was used.");
$pdfobj->addText("<PAGE WIDTH=11 HEIGHT=8.5 COLS=2 COLSPACE=1>");
$pdfobj->addText("This is on page 3 - used the PAGE comand to go to the next page.");
$pdfobj->setPre("html");
$pdf = $pdfobj->buildPDF(true,'c:\temp\hello.pdf'); # this is the output file
print $pdf->Msg . "\n"; # if all went well should return OK
print $pdf->Pages . "\n"; # this should be the page count
# note there is also a $pdf->Bytes with the raw bytes of the PDF output

# Normally you would leave the server running for the next
# request but since this is just a sample we'll stop it.
$pdfobj->stopServer();
?>
