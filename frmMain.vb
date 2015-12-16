'------------------------------------------------------------------------------------------------
' Filename    : modCommonFunctions.vb
' Purpose     : This is the common module that provides generic functions 
' Created By  : Felix Kang - I-CAT Computing (28 JUL 2005)
' Note        : 
' Assumptions : - Code is based on Visual Basic .NET (Visual Studio 2003)
'               - System.Drawing is added as reference
'               - System.Windows.Forms is added as reference
'------------------------------------------------------------------------------------------------
' History
' - 28 JUL 2005 : Creation date of the module
'------------------------------------------------------------------------------------------------

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
Friend WithEvents mnMain As System.Windows.Forms.MainMenu
Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
Friend WithEvents cmdStopMonitor As System.Windows.Forms.Button
Friend WithEvents cmdStartMonitor As System.Windows.Forms.Button
Friend WithEvents staMain As System.Windows.Forms.StatusBar
Friend WithEvents tmrMain As System.Windows.Forms.Timer
Friend WithEvents pnlInfo As System.Windows.Forms.StatusBarPanel
Friend WithEvents pnlDateTime As System.Windows.Forms.StatusBarPanel
Friend WithEvents lvwMain As System.Windows.Forms.ListView
Friend WithEvents clmEnabled As System.Windows.Forms.ColumnHeader
Friend WithEvents clmAppName As System.Windows.Forms.ColumnHeader
Friend WithEvents clmExecutables As System.Windows.Forms.ColumnHeader
Friend WithEvents clmAppPath As System.Windows.Forms.ColumnHeader
Friend WithEvents clmLastRestart As System.Windows.Forms.ColumnHeader
Friend WithEvents clmExtraInfo As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.components = New System.ComponentModel.Container
Me.mnMain = New System.Windows.Forms.MainMenu
Me.MenuItem1 = New System.Windows.Forms.MenuItem
Me.MenuItem2 = New System.Windows.Forms.MenuItem
Me.MenuItem3 = New System.Windows.Forms.MenuItem
Me.lvwMain = New System.Windows.Forms.ListView
Me.cmdStopMonitor = New System.Windows.Forms.Button
Me.cmdStartMonitor = New System.Windows.Forms.Button
Me.staMain = New System.Windows.Forms.StatusBar
Me.tmrMain = New System.Windows.Forms.Timer(Me.components)
Me.pnlInfo = New System.Windows.Forms.StatusBarPanel
Me.pnlDateTime = New System.Windows.Forms.StatusBarPanel
Me.clmEnabled = New System.Windows.Forms.ColumnHeader
Me.clmAppName = New System.Windows.Forms.ColumnHeader
Me.clmExecutables = New System.Windows.Forms.ColumnHeader
Me.clmAppPath = New System.Windows.Forms.ColumnHeader
Me.clmLastRestart = New System.Windows.Forms.ColumnHeader
Me.clmExtraInfo = New System.Windows.Forms.ColumnHeader
CType(Me.pnlInfo, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.pnlDateTime, System.ComponentModel.ISupportInitialize).BeginInit()
Me.SuspendLayout()
'
'mnMain
'
Me.mnMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
'
'MenuItem1
'
Me.MenuItem1.Index = 0
Me.MenuItem1.Text = "&Flie"
'
'MenuItem2
'
Me.MenuItem2.Index = 1
Me.MenuItem2.Text = "&Tools"
'
'MenuItem3
'
Me.MenuItem3.Index = 2
Me.MenuItem3.Text = "&Help"
'
'lvwMain
'
Me.lvwMain.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.clmEnabled, Me.clmAppName, Me.clmExecutables, Me.clmAppPath, Me.clmLastRestart, Me.clmExtraInfo})
Me.lvwMain.FullRowSelect = True
Me.lvwMain.GridLines = True
Me.lvwMain.Location = New System.Drawing.Point(5, 6)
Me.lvwMain.Name = "lvwMain"
Me.lvwMain.Size = New System.Drawing.Size(655, 349)
Me.lvwMain.TabIndex = 0
Me.lvwMain.View = System.Windows.Forms.View.Details
'
'cmdStopMonitor
'
Me.cmdStopMonitor.Location = New System.Drawing.Point(537, 365)
Me.cmdStopMonitor.Name = "cmdStopMonitor"
Me.cmdStopMonitor.Size = New System.Drawing.Size(123, 29)
Me.cmdStopMonitor.TabIndex = 1
Me.cmdStopMonitor.Text = "Stop Monitoring"
'
'cmdStartMonitor
'
Me.cmdStartMonitor.Location = New System.Drawing.Point(404, 365)
Me.cmdStartMonitor.Name = "cmdStartMonitor"
Me.cmdStartMonitor.Size = New System.Drawing.Size(123, 29)
Me.cmdStartMonitor.TabIndex = 2
Me.cmdStartMonitor.Text = "Start Monitoring"
'
'staMain
'
Me.staMain.Location = New System.Drawing.Point(0, 403)
Me.staMain.Name = "staMain"
Me.staMain.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.pnlInfo, Me.pnlDateTime})
Me.staMain.ShowPanels = True
Me.staMain.Size = New System.Drawing.Size(667, 22)
Me.staMain.TabIndex = 3
'
'tmrMain
'
Me.tmrMain.Enabled = True
Me.tmrMain.Interval = 1000
'
'pnlInfo
'
Me.pnlInfo.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
Me.pnlInfo.Width = 556
'
'pnlDateTime
'
Me.pnlDateTime.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
Me.pnlDateTime.Text = "DateAndTime"
Me.pnlDateTime.Width = 95
'
'clmEnabled
'
Me.clmEnabled.Text = "Enabled"
Me.clmEnabled.Width = 65
'
'clmAppName
'
Me.clmAppName.Text = "Application Name"
Me.clmAppName.Width = 147
'
'clmExecutables
'
Me.clmExecutables.Text = "Executable"
Me.clmExecutables.Width = 127
'
'clmAppPath
'
Me.clmAppPath.Text = "Application Path"
Me.clmAppPath.Width = 155
'
'clmLastRestart
'
Me.clmLastRestart.Text = "Last Restart"
Me.clmLastRestart.Width = 94
'
'clmExtraInfo
'
Me.clmExtraInfo.Text = "Extra Info"
Me.clmExtraInfo.Width = 120
'
'frmMain
'
Me.AutoScaleBaseSize = New System.Drawing.Size(6, 16)
Me.ClientSize = New System.Drawing.Size(667, 425)
Me.Controls.Add(Me.staMain)
Me.Controls.Add(Me.cmdStartMonitor)
Me.Controls.Add(Me.cmdStopMonitor)
Me.Controls.Add(Me.lvwMain)
Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.Menu = Me.mnMain
Me.Name = "frmMain"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "Software Restarter"
CType(Me.pnlInfo, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.pnlDateTime, System.ComponentModel.ISupportInitialize).EndInit()
Me.ResumeLayout(False)

    End Sub

#End Region

Private Sub cmdStartMonitor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStartMonitor.Click


End Sub

Private Sub Initialise()


  'Read settings from XML file

End Sub

End Class
