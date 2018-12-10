Imports Microsoft.Office.Tools.Ribbon

Public Class ShunnReviewRibbon
	Private Sub ShunnShowHide_Click(sender As Object, e As RibbonControlEventArgs) Handles ShunnShowHide.Click
		Dim btn As RibbonToggleButton = DirectCast(sender, RibbonToggleButton)
		If btn.Checked Then
			ThisAddIn.AddIn.ShowTaskPane()
		Else
			ThisAddIn.AddIn.HideTaskPane()
		End If
	End Sub

	Private Sub ShunnApply_Click(sender As Object, e As RibbonControlEventArgs) Handles ShunnApply.Click
		ThisAddIn.AddIn.ApplyFormatting()
	End Sub

	Private Sub ShunnReviewRibbon_Load(sender As Object, e As RibbonUIEventArgs) Handles Me.Load
		ShunnShowHide.Checked = False
		AddHandler ThisAddIn.AddIn.taskPane.VisibleChanged, AddressOf UpdateShowButton
	End Sub

	Private Sub UpdateShowButton()
		If ThisAddIn.AddIn.taskPane.Visible = True Then
			ShunnShowHide.Checked = True
		Else
			ShunnShowHide.Checked = False
		End If
	End Sub
End Class
