Imports System.Windows.Forms
Imports System.Drawing

Public Class SBO_Form
    Inherits System.Windows.Forms.Form
    Private Enum Images
        HeaderLeftOff = 0
        HeaderLeftOn = 1
        HeaderMinimizeOff = 2
        HeaderMinimizeOn = 3
        HeaderMaximizeOff = 4
        HeaderMaximizeOn = 5
        HeaderCloseOff = 6
        HeadercloseOn = 7
    End Enum
    Private strURL, strHeader, strCompany As String

    Private blnMoving, blnResizingX, blnResizingY As Boolean
    Private MouseDownX, MouseDownY, MouseResizeX, MouseResizeY As Integer
    Private MinWidth As Integer
    Private MinHeight As Integer

    Public Shadows Property Text() As String
        Get
            Return Me.lblHeader.Text
        End Get
        Set(ByVal Value As String)
            Me.lblHeader.Text = Value
            MyBase.Text = Value
        End Set
    End Property

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        blnMoving = False
        blnResizingX = False
        blnResizingY = False

        MinHeight = Me.Panelhead.Height
        MinWidth = Me.Width
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
    Friend WithEvents PictureBoxMinimize As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBoxMaximize As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBoxLeft As System.Windows.Forms.PictureBox
    Friend WithEvents Panelhead As System.Windows.Forms.Panel
    Friend WithEvents PictureBoxClose As System.Windows.Forms.PictureBox
    Friend WithEvents ImageListHeader As System.Windows.Forms.ImageList
    Friend WithEvents PictureBoxLeftBit As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBoxRightBit As System.Windows.Forms.PictureBox
    Friend WithEvents lblHeader As System.Windows.Forms.Label
    Friend WithEvents PictureBoxCenter As System.Windows.Forms.PictureBox
    Friend WithEvents lblBorderEast As System.Windows.Forms.Label
    Friend WithEvents lblBorderWest As System.Windows.Forms.Label
    Friend WithEvents PanelSouth As System.Windows.Forms.Panel
    Friend WithEvents lblBorderSouth As System.Windows.Forms.Label
    Friend WithEvents lblBorderSouthEast As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SBO_Form))
        Me.PictureBoxMinimize = New System.Windows.Forms.PictureBox
        Me.PictureBoxMaximize = New System.Windows.Forms.PictureBox
        Me.PictureBoxLeft = New System.Windows.Forms.PictureBox
        Me.Panelhead = New System.Windows.Forms.Panel
        Me.lblHeader = New System.Windows.Forms.Label
        Me.PictureBoxClose = New System.Windows.Forms.PictureBox
        Me.PictureBoxLeftBit = New System.Windows.Forms.PictureBox
        Me.PictureBoxRightBit = New System.Windows.Forms.PictureBox
        Me.PictureBoxCenter = New System.Windows.Forms.PictureBox
        Me.ImageListHeader = New System.Windows.Forms.ImageList(Me.components)
        Me.lblBorderWest = New System.Windows.Forms.Label
        Me.lblBorderEast = New System.Windows.Forms.Label
        Me.PanelSouth = New System.Windows.Forms.Panel
        Me.lblBorderSouthEast = New System.Windows.Forms.Label
        Me.lblBorderSouth = New System.Windows.Forms.Label
        CType(Me.PictureBoxMinimize, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxMaximize, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxLeft, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panelhead.SuspendLayout()
        CType(Me.PictureBoxClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxLeftBit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxRightBit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBoxCenter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelSouth.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBoxMinimize
        '
        Me.PictureBoxMinimize.Dock = System.Windows.Forms.DockStyle.Right
        Me.PictureBoxMinimize.Image = CType(resources.GetObject("PictureBoxMinimize.Image"), System.Drawing.Image)
        Me.PictureBoxMinimize.Location = New System.Drawing.Point(581, 0)
        Me.PictureBoxMinimize.Name = "PictureBoxMinimize"
        Me.PictureBoxMinimize.Size = New System.Drawing.Size(31, 28)
        Me.PictureBoxMinimize.TabIndex = 3
        Me.PictureBoxMinimize.TabStop = False
        '
        'PictureBoxMaximize
        '
        Me.PictureBoxMaximize.Dock = System.Windows.Forms.DockStyle.Right
        Me.PictureBoxMaximize.Image = CType(resources.GetObject("PictureBoxMaximize.Image"), System.Drawing.Image)
        Me.PictureBoxMaximize.Location = New System.Drawing.Point(612, 0)
        Me.PictureBoxMaximize.Name = "PictureBoxMaximize"
        Me.PictureBoxMaximize.Size = New System.Drawing.Size(31, 28)
        Me.PictureBoxMaximize.TabIndex = 4
        Me.PictureBoxMaximize.TabStop = False
        '
        'PictureBoxLeft
        '
        Me.PictureBoxLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.PictureBoxLeft.Image = CType(resources.GetObject("PictureBoxLeft.Image"), System.Drawing.Image)
        Me.PictureBoxLeft.Location = New System.Drawing.Point(3, 0)
        Me.PictureBoxLeft.Name = "PictureBoxLeft"
        Me.PictureBoxLeft.Size = New System.Drawing.Size(31, 28)
        Me.PictureBoxLeft.TabIndex = 2
        Me.PictureBoxLeft.TabStop = False
        '
        'Panelhead
        '
        Me.Panelhead.Controls.Add(Me.lblHeader)
        Me.Panelhead.Controls.Add(Me.PictureBoxLeft)
        Me.Panelhead.Controls.Add(Me.PictureBoxMinimize)
        Me.Panelhead.Controls.Add(Me.PictureBoxMaximize)
        Me.Panelhead.Controls.Add(Me.PictureBoxClose)
        Me.Panelhead.Controls.Add(Me.PictureBoxLeftBit)
        Me.Panelhead.Controls.Add(Me.PictureBoxRightBit)
        Me.Panelhead.Controls.Add(Me.PictureBoxCenter)
        Me.Panelhead.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panelhead.Location = New System.Drawing.Point(0, 0)
        Me.Panelhead.Name = "Panelhead"
        Me.Panelhead.Size = New System.Drawing.Size(699, 28)
        Me.Panelhead.TabIndex = 0
        '
        'lblHeader
        '
        Me.lblHeader.BackColor = System.Drawing.Color.MidnightBlue
        Me.lblHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeader.ForeColor = System.Drawing.SystemColors.Control
        Me.lblHeader.Location = New System.Drawing.Point(40, 7)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(274, 16)
        Me.lblHeader.TabIndex = 0
        Me.lblHeader.Text = "SBOFakeForm"
        '
        'PictureBoxClose
        '
        Me.PictureBoxClose.Dock = System.Windows.Forms.DockStyle.Right
        Me.PictureBoxClose.Image = CType(resources.GetObject("PictureBoxClose.Image"), System.Drawing.Image)
        Me.PictureBoxClose.Location = New System.Drawing.Point(643, 0)
        Me.PictureBoxClose.Name = "PictureBoxClose"
        Me.PictureBoxClose.Size = New System.Drawing.Size(31, 28)
        Me.PictureBoxClose.TabIndex = 4
        Me.PictureBoxClose.TabStop = False
        '
        'PictureBoxLeftBit
        '
        Me.PictureBoxLeftBit.Dock = System.Windows.Forms.DockStyle.Left
        Me.PictureBoxLeftBit.Image = CType(resources.GetObject("PictureBoxLeftBit.Image"), System.Drawing.Image)
        Me.PictureBoxLeftBit.Location = New System.Drawing.Point(0, 0)
        Me.PictureBoxLeftBit.Name = "PictureBoxLeftBit"
        Me.PictureBoxLeftBit.Size = New System.Drawing.Size(3, 28)
        Me.PictureBoxLeftBit.TabIndex = 2
        Me.PictureBoxLeftBit.TabStop = False
        '
        'PictureBoxRightBit
        '
        Me.PictureBoxRightBit.Dock = System.Windows.Forms.DockStyle.Right
        Me.PictureBoxRightBit.Image = CType(resources.GetObject("PictureBoxRightBit.Image"), System.Drawing.Image)
        Me.PictureBoxRightBit.Location = New System.Drawing.Point(674, 0)
        Me.PictureBoxRightBit.Name = "PictureBoxRightBit"
        Me.PictureBoxRightBit.Size = New System.Drawing.Size(25, 28)
        Me.PictureBoxRightBit.TabIndex = 4
        Me.PictureBoxRightBit.TabStop = False
        '
        'PictureBoxCenter
        '
        Me.PictureBoxCenter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBoxCenter.Image = CType(resources.GetObject("PictureBoxCenter.Image"), System.Drawing.Image)
        Me.PictureBoxCenter.Location = New System.Drawing.Point(0, 0)
        Me.PictureBoxCenter.Name = "PictureBoxCenter"
        Me.PictureBoxCenter.Size = New System.Drawing.Size(699, 28)
        Me.PictureBoxCenter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBoxCenter.TabIndex = 2
        Me.PictureBoxCenter.TabStop = False
        '
        'ImageListHeader
        '
        Me.ImageListHeader.ImageStream = CType(resources.GetObject("ImageListHeader.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListHeader.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageListHeader.Images.SetKeyName(0, "")
        Me.ImageListHeader.Images.SetKeyName(1, "")
        Me.ImageListHeader.Images.SetKeyName(2, "")
        Me.ImageListHeader.Images.SetKeyName(3, "")
        Me.ImageListHeader.Images.SetKeyName(4, "")
        Me.ImageListHeader.Images.SetKeyName(5, "")
        Me.ImageListHeader.Images.SetKeyName(6, "")
        Me.ImageListHeader.Images.SetKeyName(7, "")
        '
        'lblBorderWest
        '
        Me.lblBorderWest.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblBorderWest.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.lblBorderWest.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblBorderWest.Location = New System.Drawing.Point(0, 28)
        Me.lblBorderWest.Name = "lblBorderWest"
        Me.lblBorderWest.Size = New System.Drawing.Size(3, 120)
        Me.lblBorderWest.TabIndex = 1
        '
        'lblBorderEast
        '
        Me.lblBorderEast.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblBorderEast.Cursor = System.Windows.Forms.Cursors.SizeWE
        Me.lblBorderEast.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblBorderEast.Location = New System.Drawing.Point(695, 28)
        Me.lblBorderEast.Name = "lblBorderEast"
        Me.lblBorderEast.Size = New System.Drawing.Size(4, 120)
        Me.lblBorderEast.TabIndex = 3
        '
        'PanelSouth
        '
        Me.PanelSouth.Controls.Add(Me.lblBorderSouthEast)
        Me.PanelSouth.Controls.Add(Me.lblBorderSouth)
        Me.PanelSouth.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelSouth.Location = New System.Drawing.Point(3, 145)
        Me.PanelSouth.Name = "PanelSouth"
        Me.PanelSouth.Size = New System.Drawing.Size(692, 3)
        Me.PanelSouth.TabIndex = 2
        '
        'lblBorderSouthEast
        '
        Me.lblBorderSouthEast.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblBorderSouthEast.Cursor = System.Windows.Forms.Cursors.SizeNWSE
        Me.lblBorderSouthEast.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblBorderSouthEast.Location = New System.Drawing.Point(672, 0)
        Me.lblBorderSouthEast.Name = "lblBorderSouthEast"
        Me.lblBorderSouthEast.Size = New System.Drawing.Size(20, 3)
        Me.lblBorderSouthEast.TabIndex = 1
        '
        'lblBorderSouth
        '
        Me.lblBorderSouth.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblBorderSouth.Cursor = System.Windows.Forms.Cursors.SizeNS
        Me.lblBorderSouth.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblBorderSouth.Location = New System.Drawing.Point(0, 0)
        Me.lblBorderSouth.Name = "lblBorderSouth"
        Me.lblBorderSouth.Size = New System.Drawing.Size(692, 3)
        Me.lblBorderSouth.TabIndex = 0
        '
        'SBO_FakeForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(699, 148)
        Me.Controls.Add(Me.PanelSouth)
        Me.Controls.Add(Me.lblBorderWest)
        Me.Controls.Add(Me.lblBorderEast)
        Me.Controls.Add(Me.Panelhead)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SBO_FakeForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SboFakeForm"
        Me.TopMost = True
        CType(Me.PictureBoxMinimize, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxMaximize, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxLeft, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panelhead.ResumeLayout(False)
        CType(Me.PictureBoxClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxLeftBit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxRightBit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBoxCenter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelSouth.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "event handling for SAP header"

#Region "SAP button events"
    Private Sub PictureBoxMinimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxMinimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub PictureBoxLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxLeft.Click
        Me.Close()
    End Sub

    Private Sub PictureBoxMaximize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxMaximize.Click
        If Me.WindowState = FormWindowState.Normal Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If
    End Sub

    Private Sub PictureBoxClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxClose.Click
        Me.Close()
    End Sub

    Private Sub PictureBoxMinimize_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxMinimize.MouseHover
        PictureBoxMinimize.Image = ImageListHeader.Images.Item(Images.HeaderMinimizeOn)
    End Sub

    Private Sub PictureBoxMinimize_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxMinimize.MouseLeave
        PictureBoxMinimize.Image = ImageListHeader.Images.Item(Images.HeaderMinimizeOff)
    End Sub

    Private Sub PictureBoxLeft_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxLeft.MouseHover
        PictureBoxLeft.Image = ImageListHeader.Images.Item(Images.HeaderLeftOn)
    End Sub

    Private Sub PictureBoxLeft_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxLeft.MouseLeave
        PictureBoxLeft.Image = ImageListHeader.Images.Item(Images.HeaderLeftOff)
    End Sub

    Private Sub PictureBoxMaximize_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxMaximize.MouseHover
        PictureBoxMaximize.Image = ImageListHeader.Images.Item(Images.HeaderMaximizeOn)
    End Sub

    Private Sub PictureBoxMaximize_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxMaximize.MouseLeave
        PictureBoxMaximize.Image = ImageListHeader.Images.Item(Images.HeaderMaximizeOff)
    End Sub

    Private Sub PictureBoxClose_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxClose.MouseHover
        PictureBoxClose.Image = ImageListHeader.Images.Item(Images.HeadercloseOn)
    End Sub

    Private Sub PictureBoxClose_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBoxClose.MouseLeave

        If ImageListHeader.Images.Count <> 0 Then
            PictureBoxClose.Image = ImageListHeader.Images.Item(Images.HeaderCloseOff)
        End If
    End Sub
#End Region
#Region "SAP drag event"

    Private Sub lblHeader_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblHeader.MouseDown
        RememberMouseMove(e)
    End Sub

    Private Sub lblHeader_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblHeader.MouseUp
        EndMouseMove(e)
    End Sub

    Private Sub lblHeader_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblHeader.MouseMove
        CheckMovement(e)
    End Sub
    Private Sub PictureBoxCenter_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxCenter.MouseDown
        RememberMouseMove(e)
    End Sub
    Private Sub PictureBoxCenter_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxCenter.MouseMove
        CheckMovement(e)
    End Sub
    Private Sub PictureBoxCenter_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxCenter.MouseUp
        EndMouseMove(e)
    End Sub
    Private Sub PictureBoxLeftBit_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxLeftBit.MouseDown
        RememberMouseMove(e)
    End Sub

    Private Sub PictureBoxLeftBit_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxLeftBit.MouseMove
        CheckMovement(e)
    End Sub

    Private Sub PictureBoxLeftBit_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxLeftBit.MouseUp
        EndMouseMove(e)
    End Sub

    Private Sub PictureBoxRightBit_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxRightBit.MouseDown
        RememberMouseMove(e)
    End Sub

    Private Sub PictureBoxRightBit_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxRightBit.MouseMove
        CheckMovement(e)
    End Sub

    Private Sub PictureBoxRightBit_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBoxRightBit.MouseUp
        EndMouseMove(e)
    End Sub
#Region "procedures for movement"
    Private Sub RememberMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            blnMoving = True

            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub
    Private Sub EndMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            blnMoving = False
        End If
    End Sub
    Private Sub CheckMovement(ByVal e As System.Windows.Forms.MouseEventArgs)
        If blnMoving Then
            Dim Temp As Point = New Point

            Temp.X = Me.Location.X + (e.X - MouseDownX)
            Temp.Y = Me.Location.Y + (e.Y - MouseDownY)

            Me.Location = Temp
            Me.Refresh()
            Application.DoEvents()
        End If
    End Sub
#End Region
#End Region
#End Region

#Region "SAP Border Handling"

#Region "Handle East Border"
    Private Sub lblBorderEast_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderEast.MouseDown
        blnResizingY = True
        MouseResizeY = e.Y
        Application.DoEvents()
    End Sub

    Private Sub lblBorderEast_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderEast.MouseUp
        blnResizingY = False
    End Sub

    Private Sub lblBorderEast_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderEast.MouseMove
        If blnResizingY Then
            Cursor.Current = Cursors.SizeWE
            Dim Temp As Point = New Point

            Temp.X = Me.Width + (e.X - MouseResizeX)
            Temp.Y = Me.Height
            If Temp.X > MinWidth Then
                Me.Size = New Size(Temp.X, Temp.Y)
                Me.Refresh()
            End If
        End If
        Application.DoEvents()
    End Sub
    Private Sub lblBorderEast_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblBorderEast.MouseHover
        'Cursor.Current = Cursors.SizeWE
    End Sub
    Private Sub lblBorderEast_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblBorderEast.MouseLeave
        'Cursor.Current = Cursors.Arrow
    End Sub
#End Region

#Region "Handle South Border"
    Private Sub lblBorderSouth_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderSouth.MouseDown
        blnResizingY = True
        MouseResizeY = e.Y
    End Sub

    Private Sub lblBorderSouth_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderSouth.MouseUp
        blnResizingY = False
    End Sub

    Private Sub lblBorderSouth_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderSouth.MouseMove
        If blnResizingY Then
            Cursor.Current = Cursors.SizeNS
            Dim Temp As Point = New Point
            Temp.X = Me.Width
            Temp.Y = Me.Height + e.Y - MouseResizeY
            If Temp.Y > MinHeight Then
                Me.Size = New Size(Temp.X, Temp.Y)
                Me.Refresh()
            End If
        End If
        Application.DoEvents()
    End Sub

#End Region

#Region "Handle SouthEast Border"
    Private Sub lblBorderSouthEast_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderSouthEast.MouseDown
        blnResizingY = True
        blnResizingX = True
    End Sub

    Private Sub lblBorderSouthEast_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderSouthEast.MouseUp
        blnResizingY = False
        blnResizingX = False
    End Sub

    Private Sub lblBorderSouthEast_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblBorderSouthEast.MouseMove
        If blnResizingY = True _
          And blnResizingX = True Then

            Dim Temp As Point = New Point
            Temp.X = Me.Width + e.X - MouseResizeX
            Temp.Y = Me.Height + e.Y - MouseResizeY
            If Temp.Y > MinHeight And Temp.X > MinWidth Then
                Me.Size = New Size(Temp.X, Temp.Y)
                Me.Refresh()
            End If
        End If
        Application.DoEvents()
    End Sub
#End Region
#End Region



End Class
