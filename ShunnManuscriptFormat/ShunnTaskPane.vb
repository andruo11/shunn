Imports System.Configuration

Public Class ShunnTaskPane
	Dim config As Configuration
	Dim configLocation As String
	Private Sub ApplyFormatting_Click(sender As Object, e As EventArgs) Handles ApplyFormatting.Click
		ThisAddIn.AddIn.ApplyFormatting()
	End Sub

	Private Sub TextBox_TextChanged(sender As Object, e As EventArgs)
		Dim textBox As Windows.Forms.TextBox = DirectCast(sender, Windows.Forms.TextBox)
		ShunnManuscriptFormat.MySettings.Default(textBox.Name) = textBox.Text
		ShunnManuscriptFormat.MySettings.Default.Save()
	End Sub
	Private Sub ShunnTaskPane_Load(sender As Object, e As EventArgs) Handles Me.Load
		Dim textBoxes() As Windows.Forms.TextBox = New Windows.Forms.TextBox() {FirstName, LastName, AuthorByline, AddressLine1,
			AddressLine2, Telephone, EmailAddress, Note, StoryTitle, TitleKeyword}
		For Each textBox As Windows.Forms.TextBox In textBoxes
			textBox.Text = ShunnManuscriptFormat.MySettings.Default(textBox.Name)
			AddHandler textBox.TextChanged, AddressOf TextBox_TextChanged
		Next
	End Sub

	Private Sub LinkLabel1_LinkClicked(sender As Object, e As Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
		System.Diagnostics.Process.Start("https://www.shunn.net/format/story.html")
	End Sub
End Class
