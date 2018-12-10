Imports Microsoft.Office.Interop
Imports System.Configuration
Imports System.Collections.Specialized

'manuscript format described at https://www.shunn.net/format/story.html
'setup projects for windows installer according to (minus launch conditions) https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-studio-2010/ff937654(v=msdn.10)
'by using visual studio 2017 installer extension at https://marketplace.visualstudio.com/items?itemName=VisualStudioClient.MicrosoftVisualStudio2017InstallerProjects
Public Class ThisAddIn
	Friend Shared AddIn As ThisAddIn
	Dim stp As ShunnTaskPane
	Friend taskPane As Microsoft.Office.Tools.CustomTaskPane
	Private Sub ThisAddIn_Startup() Handles Me.Startup
		ThisAddIn.AddIn = Me

		stp = New ShunnTaskPane
		taskPane = Me.CustomTaskPanes.Add(stp, "Shunn manuscript format")
		taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
		taskPane.Width = 750
		taskPane.Visible = False

	End Sub
	Friend Sub ApplyFormatting()

		Dim doc As Word.Document = Application.ActiveDocument
		Dim rng As Word.Range = Application.Selection.Range
		If rng.End <= rng.Start Then
			MsgBox("Select the body of your story and then click this button.")
			Exit Sub
		End If
		rng.Copy()


		Dim wordCount As Integer = rng.ComputeStatistics(Word.WdStatistic.wdStatisticWords)
		If wordCount < 7000 Then 'less than novella length, round to nearest 100 words
			wordCount = Math.Round(CDbl(wordCount) / 100.0) * 100
		Else 'round to nearest 500 words
			wordCount = Math.Round(CDbl(wordCount) / 500.0) * 500
		End If

		Dim doc2 As Word.Document = MakeBlankDocument()
		Dim header1Lines As Integer = FormatFirstHeader(doc2, wordCount)
		FormatPrimaryHeader(doc2)

		doc2.Range.Paste()
		Dim oldLength As Integer = doc2.Range.End - doc2.Range.Start

		'add story title and byline halfway down the first page
		With doc2.Range
			.InsertParagraphBefore()
			.InsertParagraphBefore()
			If stp.AuthorByline.Text.Trim <> "" Then
				.InsertBefore("by " & stp.AuthorByline.Text)
			Else
				.InsertBefore("by " & stp.FirstName.Text & " " & stp.LastName.Text)
			End If
			.InsertParagraphBefore()
			.InsertBefore(stp.StoryTitle.Text)
			For x As Integer = 1 To 11 - header1Lines
				.InsertParagraphBefore()
			Next
		End With

		Dim newLength As Integer = doc2.Range.End - doc2.Range.Start
		rng = doc2.Range
		For Each para As Word.Paragraph In rng.Paragraphs
			para.LineSpacing = 24
			para.SpaceAfter = 0
			para.SpaceBefore = 0
		Next

		'format only the byline and story title
		rng.SetRange(1, newLength - oldLength)
		rng.ParagraphFormat.Style = Word.WdBuiltinStyle.wdStyleNormal
		rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

		rng = doc2.Range
		FormatRange(rng)

		'format only the body text after the byline and story title
		rng.SetRange(newLength - oldLength, newLength)
		rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
		For Each para As Word.Paragraph In rng.Paragraphs
			If para.Style.NameLocal.StartsWith("Head") Then
				para.FirstLineIndent = 0
				para.Range.Font.Bold = True
			Else
				para.FirstLineIndent = 36
			End If
		Next
		For Each letter As Word.Range In rng.Characters
			If letter.Italic Then letter.Underline = True
		Next
		MakeReplacements(rng)
		rng.Collapse(Word.WdCollapseDirection.wdCollapseStart)
		rng.Select()
	End Sub
	Private Function MakeBlankDocument() As Word.Document
		Dim doc As Word.Document = Application.Documents.Add()
		With doc.PageSetup
			.PageWidth = 612
			.PageHeight = 792
			.BottomMargin = 72
			.TopMargin = 72
			.LeftMargin = 72
			.RightMargin = 72
			.DifferentFirstPageHeaderFooter = True
			.HeaderDistance = 72
		End With
		Return (doc)
	End Function
	Friend Sub FormatRange(ByRef rng As Word.Range)
		With rng.Font
			.Name = "Courier"
			.Size = 12
			.Color = Word.WdColor.wdColorAutomatic
		End With
	End Sub
	Friend Function FormatFirstHeader(ByRef doc As Word.Document, ByVal wordCount As Integer) As Integer
		Dim h1Obj As Word.HeaderFooter = doc.Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage)
		Dim header1 As Word.Range = h1Obj.Range
		Dim retVal As Integer
		With header1
			.Text = "about " & wordCount.ToString("N0") & " words"
			.Collapse(Word.WdCollapseDirection.wdCollapseStart)
			.InsertAlignmentTab(Word.WdAlignmentTabAlignment.wdRight, Word.WdAlignmentTabRelative.wdMargin)
			.InsertBefore(stp.FirstName.Text & " " & stp.LastName.Text)
		End With
		With h1Obj.Range
			.InsertParagraphAfter()
			.InsertAfter(stp.AddressLine1.Text)
			.InsertParagraphAfter()
			.InsertAfter(stp.AddressLine2.Text)
			.InsertParagraphAfter()
			.InsertAfter(stp.Telephone.Text)
			.InsertParagraphAfter()
			.InsertAfter(stp.EmailAddress.Text)
			If stp.Note.Text.Trim <> "" Then
				.InsertParagraphAfter()
				.InsertParagraphAfter()
				.InsertAfter(stp.Note.Text)
				retVal = 1
			Else
				retVal = 0
			End If
		End With
		FormatRange(h1Obj.Range)
		Return (retVal)
	End Function
	Friend Sub FormatPrimaryHeader(ByRef doc As Word.Document)
		Dim h2Obj As Word.HeaderFooter = doc.Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)
		Dim header2 As Word.Range = h2Obj.Range
		header2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
		header2.Text = stp.LastName.Text & " / " & stp.TitleKeyword.Text & " /"
		h2Obj.PageNumbers.Add(FirstPage:=False)
		h2Obj.Range.InsertParagraphAfter()
		FormatRange(h2Obj.Range)
	End Sub
	Private Sub MakeReplacements(ByRef rng As Word.Range)
		rng.Select()
		With Application.Selection
			With .Find
				.ClearFormatting()
				.Replacement.ClearFormatting()
				.Text = "—"
				.Replacement.Text = "--"
				.Forward = True
				.Wrap = Word.WdFindWrap.wdFindContinue
				.MatchWildcards = False
				.Execute(Replace:=Word.WdReplace.wdReplaceAll)
			End With
		End With
	End Sub
	Friend Sub ShowTaskPane()
		taskPane.Visible = True
	End Sub
	Friend Sub HideTaskPane()
		taskPane.Visible = False
	End Sub
End Class