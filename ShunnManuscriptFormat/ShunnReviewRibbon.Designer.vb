Partial Class ShunnReviewRibbon
	Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

	<System.Diagnostics.DebuggerNonUserCode()> _
	Public Sub New(ByVal container As System.ComponentModel.IContainer)
		MyClass.New()

		'Required for Windows.Forms Class Composition Designer support
		If (container IsNot Nothing) Then
			container.Add(Me)
		End If

	End Sub

	<System.Diagnostics.DebuggerNonUserCode()> _
	Public Sub New()
		MyBase.New(Globals.Factory.GetRibbonFactory())

		'This call is required by the Component Designer.
		InitializeComponent()

	End Sub

	'Component overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

	'Required by the Component Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Component Designer
	'It can be modified using the Component Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
		Me.Tab1 = Me.Factory.CreateRibbonTab
		Me.Group1 = Me.Factory.CreateRibbonGroup
		Me.Button1 = Me.Factory.CreateRibbonButton
		Me.ShunnShowHide = Me.Factory.CreateRibbonToggleButton
		Me.ShunnApply = Me.Factory.CreateRibbonButton
		Me.ShunnShow = Me.Factory.CreateRibbonToggleButton
		Me.Tab1.SuspendLayout()
		Me.Group1.SuspendLayout()
		Me.SuspendLayout()
		'
		'Tab1
		'
		Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
		Me.Tab1.ControlId.OfficeId = "TabReviewWord"
		Me.Tab1.Groups.Add(Me.Group1)
		Me.Tab1.Label = "TabReviewWord"
		Me.Tab1.Name = "Tab1"
		'
		'Group1
		'
		Me.Group1.Items.Add(Me.ShunnShowHide)
		Me.Group1.Items.Add(Me.ShunnApply)
		Me.Group1.Label = "Shunn format"
		Me.Group1.Name = "Group1"
		'
		'Button1
		'
		Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
		Me.Button1.Label = "Apply formatting"
		Me.Button1.Name = "Button1"
		Me.Button1.ShowImage = True
		'
		'ShunnShowHide
		'
		Me.ShunnShowHide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
		Me.ShunnShowHide.Image = Global.ShunnManuscriptFormat.My.Resources.Resources.baseline_account_box_black_18dp
		Me.ShunnShowHide.Label = "Show/hide user info"
		Me.ShunnShowHide.Name = "ShunnShowHide"
		Me.ShunnShowHide.ShowImage = True
		'
		'ShunnApply
		'
		Me.ShunnApply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
		Me.ShunnApply.Image = Global.ShunnManuscriptFormat.My.Resources.Resources.Book_Open2
		Me.ShunnApply.Label = "Create document from selection"
		Me.ShunnApply.Name = "ShunnApply"
		Me.ShunnApply.ShowImage = True
		'
		'ShunnShow
		'
		Me.ShunnShow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
		Me.ShunnShow.Image = Global.ShunnManuscriptFormat.My.Resources.Resources.baseline_account_box_black_18dp
		Me.ShunnShow.Label = "Show/hide form"
		Me.ShunnShow.Name = "ShunnShow"
		Me.ShunnShow.ShowImage = True
		'
		'ShunnReviewRibbon
		'
		Me.Name = "ShunnReviewRibbon"
		Me.RibbonType = "Microsoft.Word.Document"
		Me.Tabs.Add(Me.Tab1)
		Me.Tab1.ResumeLayout(False)
		Me.Tab1.PerformLayout()
		Me.Group1.ResumeLayout(False)
		Me.Group1.PerformLayout()
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
	Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
	Friend WithEvents ShunnApply As Microsoft.Office.Tools.Ribbon.RibbonButton
	Friend WithEvents ShunnShowHide As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
	Friend WithEvents ShunnShow As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
	Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

	<System.Diagnostics.DebuggerNonUserCode()> _
	Friend ReadOnly Property ShunnReviewRibbon() As ShunnReviewRibbon
		Get
			Return Me.GetRibbon(Of ShunnReviewRibbon)()
		End Get
	End Property
End Class
