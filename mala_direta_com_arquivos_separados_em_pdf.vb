Sub QuebraNaSeccao()
    Dim Arquivo As Integer
    Dim CaminhoArquivo As String
    Dim TextoProximaLinha As String
    
	'Set reading the file that contains the names of files that will be generated.
    Arquivo = FreeFile
    CaminhoArquivo = "F:\Documentos\Certificados\participantes.txt"
 
    'Open file for reading.
    Open CaminhoArquivo For Input As Arquivo

   ' Used to set criteria for moving through the document by section.
   Application.Browser.Target = wdBrowseSection

	'A mail merge document ends with a section break next page.
	'Subtracting one from the section count stop error message.
	For i = 1 To ((ActiveDocument.Sections.Count) - 1)   
		'Note: If a document does not end with a section break,
		'substitute the following line of code for the one above:
		'For I = 1 To ActiveDocument.Sections.Count

		'Select and copy the section text to the clipboard.
		ActiveDocument.Bookmarks("\Section").Range.Copy

		'Create a new document to paste text from clipboard.
		Documents.Add
		Selection.Paste
		  
		'Altera a orientação da página para paisagem
		Orientation
		'Deletes the last page (use only if necessary)
		DeleteLastLine
      
		'Removes the break that is copied at the end of the section, if any.
		Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
		Selection.Delete Unit:=wdCharacter, Count:=1
		ChangeFileOpenDirectory "F:\Documentos\Certificados Flisol 2016\Certificados\"
     
		'It makes the line reading
		Line Input #Arquivo, TextoProximaLinha
        TextoProximaLinha = TextoProximaLinha
     
		'Export to .pdf and customize the file name to the line that was read
		 ActiveDocument.ExportAsFixedFormat OutputFileName:= _
		"F:\Documentos\Certificados Flisol 2016\Certificados\" & TextoProximaLinha & ".pdf" _
		, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
		wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
		Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
		CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
		BitmapMissingFonts:=True, UseISO19005_1:=False
     
		'Closes the "temporary" file from Word without saving changes
		ActiveDocument.Close savechanges:=wdDoNotSaveChanges
		'Move the selection to the next section in the document.
		Application.Browser.Next
   Next i
   ActiveDocument.Close savechanges:=wdDoNotSaveChanges
End Sub

Sub Orientation()
	'If the page orientation is portrait in it is changed to landscape
	'This is a particular case in issuing certificates. Make sure that in your case there is a need
    If Selection.PageSetup.Orientation = wdOrientPortrait Then
        Selection.PageSetup.Orientation = wdOrientLandscape
    Else
        Selection.PageSetup.Orientation = wdOrientPortrait
    End If
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
End Sub
	
Sub DeleteLastLine()
	'This is a particular case in issuing certificates. Make sure that in your case there is a need
	Selection.HomeKey Unit:=wdStory
    Selection.EndKey Unit:=wdStory
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub

