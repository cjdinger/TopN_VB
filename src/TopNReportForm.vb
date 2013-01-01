Imports System.Windows
Imports SAS.Shared.AddIns

Namespace SASPress.CustomTasks.TopNVB

    Public Class TopNReportForm
        Inherits System.Windows.Forms.Form

        ' Properties
        Public Property Model() As TopNReport
            Get
                Return Me._model
            End Get
            Set(ByVal value As TopNReport)
                Me._model = value
            End Set
        End Property

        Private hashFormats As Hashtable = New Hashtable
        Private noCategory As String = "<No category>"
        Private selectMeasurePrompt As String = "<Select a measure>"
        Private selectReportPrompt As String = "<Select a report column>"

        ' Fields
        Private _model As TopNReport = Nothing

        Private Sub CommitChangesToModel()
            Me.Model.N = Convert.ToInt32(Me.nudN.Value)
            Me.Model.IncludeChart = Me.chkIncludeChart.Checked
            Me.Model.Statistic = DirectCast([Enum].Parse(GetType(TopNReport.eStatistic), Me.cmbStatistic.SelectedItem.ToString, False), TopNReport.eStatistic)
            Me.Model.ReportColumn = Me.cmbReport.SelectedItem.ToString
            Me.Model.MeasureColumn = Me.cmbMeasure.SelectedItem.ToString
            Me.Model.MeasureFormat = CStr(Me.hashFormats.Item(Me.Model.MeasureColumn))
            If (Me.cmbCategory.SelectedItem.ToString = Me.noCategory) Then
                Me.Model.CategoryColumn = ""
            Else
                Me.Model.CategoryColumn = Me.cmbCategory.SelectedItem.ToString
            End If
            Me.Model.Title = Me.txtTitle.Text
            Me.Model.Footnote = Me.txtFootnote.Text
        End Sub

        Private Sub LoadColumns()
            Me.cmbReport.Items.Add(Me.selectReportPrompt)
            Me.cmbMeasure.Items.Add(Me.selectMeasurePrompt)
            Try
                Dim accessor As ISASTaskDataAccessor = Me.Model.Consumer.ActiveData.Accessor
                If accessor.Open Then
                    Dim i As Integer
                    For i = 0 To accessor.ColumnCount - 1
                        Dim column As ISASTaskDataColumn = accessor.ColumnInfoByIndex(i)
                        If ((column.Group = VariableGroup.Character) OrElse (column.Group = VariableGroup.Date)) Then
                            Me.cmbReport.Items.Add(column.Name)
                            Me.cmbCategory.Items.Add(column.Name)
                        Else
                            Me.cmbMeasure.Items.Add(column.Name)
                            Me.hashFormats.Add(column.Name, column.Format)
                        End If
                    Next i
                    Dim j As Integer
                    For j = 0 To accessor.ColumnCount - 1
                        Dim column2 As ISASTaskDataColumn = accessor.ColumnInfoByIndex(j)
                        If (column2.Group = VariableGroup.Numeric) Then
                            Me.cmbReport.Items.Add(column2.Name)
                            Me.cmbCategory.Items.Add(column2.Name)
                        End If
                    Next j
                    Me.cmbCategory.Items.Add(Me.noCategory)
                    accessor.Close()
                Else
                    Dim str As String = "UNKNOWN"
                    If (Not Me.Model.Consumer.ActiveData Is Nothing) Then
                        str = String.Format("{0}.{1}", Me.Model.Consumer.ActiveData.Library, Me.Model.Consumer.ActiveData.Member)
                    End If
                    System.Windows.Forms.MessageBox.Show(String.Format("ERROR: Could not read column information from data {0}.", str))
                    MyBase.Close()
                End If
            Catch
            End Try
        End Sub

        Private Sub LoadSettingsFromModel()
            If ((Me.Model.ReportColumn.Length > 0) AndAlso (Me.cmbReport.FindStringExact(Me.Model.ReportColumn) <> -1)) Then
                Me.cmbReport.SelectedItem = Me.Model.ReportColumn
            Else
                Me.cmbReport.SelectedItem = Me.selectReportPrompt
            End If
            If ((Me.Model.CategoryColumn.Length > 0) AndAlso (Me.cmbCategory.FindStringExact(Me.Model.CategoryColumn) <> -1)) Then
                Me.cmbCategory.SelectedItem = Me.Model.CategoryColumn
            Else
                Me.cmbCategory.SelectedItem = Me.noCategory
            End If
            If ((Me.Model.MeasureColumn.Length > 0) AndAlso (Me.cmbMeasure.FindStringExact(Me.Model.MeasureColumn) <> -1)) Then
                Me.cmbMeasure.SelectedItem = Me.Model.MeasureColumn
            Else
                Me.cmbMeasure.SelectedItem = Me.selectMeasurePrompt
            End If
            Me.cmbStatistic.SelectedItem = Me.Model.Statistic.ToString
            Me.nudN.Value = Me.Model.N
            Me.chkIncludeChart.Checked = Me.Model.IncludeChart
            Me.txtTitle.Text = Me.Model.Title
            Me.txtFootnote.Text = Me.Model.Footnote
        End Sub

        Private Sub LoadStatistics()
            Me.cmbStatistic.Items.Add(TopNReport.eStatistic.Sum.ToString)
            Me.cmbStatistic.Items.Add(TopNReport.eStatistic.Average.ToString)
            Me.cmbStatistic.Items.Add(TopNReport.eStatistic.Count.ToString)
        End Sub

        Protected Overrides Sub OnClosed(ByVal e As EventArgs)
            If (DialogResult.OK = MyBase.DialogResult) Then
                Me.CommitChangesToModel()
            End If
            MyBase.OnClosed(e)
        End Sub

        Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
            If ((MyBase.DialogResult = DialogResult.OK) AndAlso Not Me.Model.Consumer.VerifyTaskClosing(Me)) Then
                MyBase.DialogResult = DialogResult.None
                e.Cancel = True
            End If
            MyBase.OnClosing(e)
        End Sub

        Protected Overrides Sub OnLoad(ByVal e As EventArgs)
            MyBase.OnLoad(e)
            Me.nudN.Minimum = 1
            If (Not Me.Model Is Nothing) Then
                Me.LoadColumns()
                Me.LoadStatistics()
                Me.LoadSettingsFromModel()
            End If
            AddHandler Me.cmbMeasure.SelectedIndexChanged, New EventHandler(AddressOf Me.OnSelectionChanged)
            AddHandler Me.cmbCategory.SelectedIndexChanged, New EventHandler(AddressOf Me.OnSelectionChanged)
            AddHandler Me.cmbStatistic.SelectedIndexChanged, New EventHandler(AddressOf Me.OnSelectionChanged)
            AddHandler Me.cmbReport.SelectedIndexChanged, New EventHandler(AddressOf Me.OnSelectionChanged)
            Me.UpdateControls()
        End Sub

        Private Sub OnSelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me.UpdateControls()
        End Sub

        Private Sub UpdateControls()
            Dim flag As Boolean = (Me.cmbStatistic.SelectedItem.ToString = TopNReport.eStatistic.Count.ToString)
            Me.cmbMeasure.Enabled = Not flag
            If ((Me.cmbReport.SelectedItem.ToString = Me.selectReportPrompt) OrElse ((Me.cmbMeasure.SelectedItem.ToString = Me.selectMeasurePrompt) AndAlso Not flag)) Then
                Me.btnOK.Enabled = False
            Else
                Me.btnOK.Enabled = True
            End If
        End Sub

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
        Private WithEvents btnOK As System.Windows.Forms.Button
        Friend WithEvents label2 As System.Windows.Forms.Label
        Friend WithEvents txtFootnote As System.Windows.Forms.TextBox
        Friend WithEvents label1 As System.Windows.Forms.Label
        Friend WithEvents txtTitle As System.Windows.Forms.TextBox
        Friend WithEvents chkIncludeChart As System.Windows.Forms.CheckBox
        Friend WithEvents nudN As System.Windows.Forms.NumericUpDown
        Friend WithEvents cmbCategory As System.Windows.Forms.ComboBox
        Friend WithEvents cmbStatistic As System.Windows.Forms.ComboBox
        Friend WithEvents cmbMeasure As System.Windows.Forms.ComboBox
        Friend WithEvents lblCategory As System.Windows.Forms.Label
        Friend WithEvents lblStatistic As System.Windows.Forms.Label
        Friend WithEvents lblMeasure As System.Windows.Forms.Label
        Friend WithEvents cmbReport As System.Windows.Forms.ComboBox
        Friend WithEvents lbReport As System.Windows.Forms.Label
        Friend WithEvents lblN As System.Windows.Forms.Label
        Private WithEvents btnCancel As System.Windows.Forms.Button
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.btnOK = New System.Windows.Forms.Button
            Me.btnCancel = New System.Windows.Forms.Button
            Me.label2 = New System.Windows.Forms.Label
            Me.txtFootnote = New System.Windows.Forms.TextBox
            Me.label1 = New System.Windows.Forms.Label
            Me.txtTitle = New System.Windows.Forms.TextBox
            Me.chkIncludeChart = New System.Windows.Forms.CheckBox
            Me.nudN = New System.Windows.Forms.NumericUpDown
            Me.cmbCategory = New System.Windows.Forms.ComboBox
            Me.cmbStatistic = New System.Windows.Forms.ComboBox
            Me.cmbMeasure = New System.Windows.Forms.ComboBox
            Me.lblCategory = New System.Windows.Forms.Label
            Me.lblStatistic = New System.Windows.Forms.Label
            Me.lblMeasure = New System.Windows.Forms.Label
            Me.cmbReport = New System.Windows.Forms.ComboBox
            Me.lbReport = New System.Windows.Forms.Label
            Me.lblN = New System.Windows.Forms.Label
            CType(Me.nudN, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'btnOK
            '
            Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.btnOK.Location = New System.Drawing.Point(360, 353)
            Me.btnOK.Name = "btnOK"
            Me.btnOK.TabIndex = 0
            Me.btnOK.Text = "OK"
            '
            'btnCancel
            '
            Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.btnCancel.Location = New System.Drawing.Point(443, 353)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.TabIndex = 1
            Me.btnCancel.Text = "Cancel"
            '
            'label2
            '
            Me.label2.Location = New System.Drawing.Point(16, 288)
            Me.label2.Name = "label2"
            Me.label2.Size = New System.Drawing.Size(406, 20)
            Me.label2.TabIndex = 30
            Me.label2.Text = "Enter a report footnote (optional):"
            Me.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'txtFootnote
            '
            Me.txtFootnote.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.txtFootnote.Location = New System.Drawing.Point(16, 312)
            Me.txtFootnote.Name = "txtFootnote"
            Me.txtFootnote.Size = New System.Drawing.Size(452, 20)
            Me.txtFootnote.TabIndex = 31
            Me.txtFootnote.Text = ""
            '
            'label1
            '
            Me.label1.Location = New System.Drawing.Point(16, 240)
            Me.label1.Name = "label1"
            Me.label1.Size = New System.Drawing.Size(406, 20)
            Me.label1.TabIndex = 28
            Me.label1.Text = "Enter a report title (optional):"
            Me.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'txtTitle
            '
            Me.txtTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.txtTitle.Location = New System.Drawing.Point(16, 264)
            Me.txtTitle.Name = "txtTitle"
            Me.txtTitle.Size = New System.Drawing.Size(452, 20)
            Me.txtTitle.TabIndex = 29
            Me.txtTitle.Text = ""
            '
            'chkIncludeChart
            '
            Me.chkIncludeChart.Location = New System.Drawing.Point(16, 200)
            Me.chkIncludeChart.Name = "chkIncludeChart"
            Me.chkIncludeChart.Size = New System.Drawing.Size(213, 21)
            Me.chkIncludeChart.TabIndex = 27
            Me.chkIncludeChart.Text = "Include chart in report"
            '
            'nudN
            '
            Me.nudN.Location = New System.Drawing.Point(192, 160)
            Me.nudN.Name = "nudN"
            Me.nudN.Size = New System.Drawing.Size(47, 20)
            Me.nudN.TabIndex = 26
            '
            'cmbCategory
            '
            Me.cmbCategory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.cmbCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.cmbCategory.Location = New System.Drawing.Point(192, 120)
            Me.cmbCategory.Name = "cmbCategory"
            Me.cmbCategory.Size = New System.Drawing.Size(279, 21)
            Me.cmbCategory.TabIndex = 24
            '
            'cmbStatistic
            '
            Me.cmbStatistic.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.cmbStatistic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.cmbStatistic.Location = New System.Drawing.Point(192, 88)
            Me.cmbStatistic.Name = "cmbStatistic"
            Me.cmbStatistic.Size = New System.Drawing.Size(279, 21)
            Me.cmbStatistic.TabIndex = 22
            '
            'cmbMeasure
            '
            Me.cmbMeasure.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.cmbMeasure.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.cmbMeasure.Location = New System.Drawing.Point(192, 56)
            Me.cmbMeasure.Name = "cmbMeasure"
            Me.cmbMeasure.Size = New System.Drawing.Size(279, 21)
            Me.cmbMeasure.TabIndex = 20
            '
            'lblCategory
            '
            Me.lblCategory.Location = New System.Drawing.Point(16, 120)
            Me.lblCategory.Name = "lblCategory"
            Me.lblCategory.Size = New System.Drawing.Size(166, 20)
            Me.lblCategory.TabIndex = 23
            Me.lblCategory.Text = "Report across category?"
            Me.lblCategory.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'lblStatistic
            '
            Me.lblStatistic.Location = New System.Drawing.Point(16, 88)
            Me.lblStatistic.Name = "lblStatistic"
            Me.lblStatistic.Size = New System.Drawing.Size(166, 20)
            Me.lblStatistic.TabIndex = 21
            Me.lblStatistic.Text = "Apply which statistic?"
            Me.lblStatistic.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'lblMeasure
            '
            Me.lblMeasure.Location = New System.Drawing.Point(16, 56)
            Me.lblMeasure.Name = "lblMeasure"
            Me.lblMeasure.Size = New System.Drawing.Size(166, 19)
            Me.lblMeasure.TabIndex = 19
            Me.lblMeasure.Text = "Which column to measure?"
            Me.lblMeasure.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'cmbReport
            '
            Me.cmbReport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.cmbReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.cmbReport.Location = New System.Drawing.Point(192, 16)
            Me.cmbReport.Name = "cmbReport"
            Me.cmbReport.Size = New System.Drawing.Size(279, 21)
            Me.cmbReport.TabIndex = 18
            '
            'lbReport
            '
            Me.lbReport.Location = New System.Drawing.Point(16, 16)
            Me.lbReport.Name = "lbReport"
            Me.lbReport.Size = New System.Drawing.Size(166, 20)
            Me.lbReport.TabIndex = 17
            Me.lbReport.Text = "Report on which column?"
            Me.lbReport.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'lblN
            '
            Me.lblN.Location = New System.Drawing.Point(16, 160)
            Me.lblN.Name = "lblN"
            Me.lblN.Size = New System.Drawing.Size(166, 19)
            Me.lblN.TabIndex = 25
            Me.lblN.Text = "Include how many top values?"
            Me.lblN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'TopNReportForm
            '
            Me.AcceptButton = Me.btnOK
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.CancelButton = Me.btnCancel
            Me.ClientSize = New System.Drawing.Size(528, 385)
            Me.Controls.Add(Me.label2)
            Me.Controls.Add(Me.txtFootnote)
            Me.Controls.Add(Me.label1)
            Me.Controls.Add(Me.txtTitle)
            Me.Controls.Add(Me.chkIncludeChart)
            Me.Controls.Add(Me.nudN)
            Me.Controls.Add(Me.cmbCategory)
            Me.Controls.Add(Me.cmbStatistic)
            Me.Controls.Add(Me.cmbMeasure)
            Me.Controls.Add(Me.lblCategory)
            Me.Controls.Add(Me.lblStatistic)
            Me.Controls.Add(Me.lblMeasure)
            Me.Controls.Add(Me.cmbReport)
            Me.Controls.Add(Me.lbReport)
            Me.Controls.Add(Me.lblN)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnOK)
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "TopNReportForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "SAS Custom Task"
            CType(Me.nudN, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

#End Region

    End Class

End Namespace