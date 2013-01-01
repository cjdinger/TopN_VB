Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Xml.Serialization
Imports System.Globalization

Imports SAS.Shared.AddIns

' Note that the namespace does not begin with "SAS."  There is a problem in the Visual Basic .NET environment
' that causes a conflict and masks out the SAS.Shared.AddIns namespace, which we need to use as well.
' Usually, the first part of the namespace contains the company name.
Namespace SASPress.CustomTasks.TopNVB

    Public Class TopNReport
        Implements SAS.Shared.AddIns.ISASTask, SAS.Shared.AddIns.ISASTaskAddIn, SAS.Shared.AddIns.ISASTaskDescription

        ' keep a reference to the ISASTaskConsumer interface for the application
        Private _consumer As ISASTaskConsumer

        Public ReadOnly Property Consumer() As ISASTaskConsumer
            Get
                Return Me._consumer
            End Get
        End Property

        Private _label As String

        Public Enum eStatistic
            ' Fields
            Average = 1
            Count = 2
            Sum = 0
        End Enum

        Private _footnote As String = ""
        Private _includeChart As Boolean = True
        Private _categoryColumn As String = ""
        Private _measureColumn As String = ""
        Private _measureFormat As String = ""
        Private _n As Integer = 10
        Private _reportColumn As String = ""
        Private _statistic As eStatistic = eStatistic.Sum
        Private _title As String = ""

        Friend Property Title() As String
            Get
                Return Me._title
            End Get
            Set(ByVal value As String)
                Me._title = value
            End Set
        End Property

        Friend Property Statistic() As eStatistic
            Get
                Return Me._statistic
            End Get
            Set(ByVal value As eStatistic)
                Me._statistic = value
            End Set
        End Property

        Friend Property CategoryColumn() As String
            Get
                Return Me._categoryColumn
            End Get
            Set(ByVal value As String)
                Me._categoryColumn = value
            End Set
        End Property
        Friend Property Footnote() As String
            Get
                Return Me._footnote
            End Get
            Set(ByVal value As String)
                Me._footnote = value
            End Set
        End Property
        Friend Property IncludeChart() As Boolean
            Get
                Return Me._includeChart
            End Get
            Set(ByVal value As Boolean)
                Me._includeChart = value
            End Set
        End Property

        Friend Property MeasureColumn() As String
            Get
                Return Me._measureColumn
            End Get
            Set(ByVal value As String)
                Me._measureColumn = value
            End Set
        End Property

        Friend Property MeasureFormat() As String
            Get
                Return Me._measureFormat
            End Get
            Set(ByVal value As String)
                Me._measureFormat = value
            End Set
        End Property

        Friend Property N() As Integer
            Get
                Return Me._n
            End Get
            Set(ByVal value As Integer)
                If (Me._n <= 0) Then
                    Throw New ArgumentOutOfRangeException("The value for N must be greater than zero.")
                End If
                Me._n = value
            End Set
        End Property
        Friend Property ReportColumn() As String
            Get
                Return Me._reportColumn
            End Get
            Set(ByVal value As String)
                Me._reportColumn = value
            End Set
        End Property

#Region "ISASTask implementation"

        ' perform any initialization for this task, and return True
        Public Function Initialize() As Boolean Implements SAS.Shared.AddIns.ISASTask.Initialize
            Return True
        End Function

        Public Property Label() As String Implements SAS.Shared.AddIns.ISASTask.Label
            Get
                Label = Me._label
            End Get
            Set(ByVal Value As String)
                Me._label = Value
            End Set
        End Property

        ' if the task creates output data, indicate it here
        Public ReadOnly Property OutputDataCount() As Integer Implements SAS.Shared.AddIns.ISASTask.OutputDataCount
            Get
                OutputDataCount = 0
            End Get
        End Property

        Public Function OutputDataInfo(ByVal Index As Integer, ByRef Source As String, ByRef Label As String) As SAS.Shared.AddIns.OutputData Implements SAS.Shared.AddIns.ISASTask.OutputDataInfo
            Source = Nothing
            Label = Nothing
            OutputDataInfo = OutputData.Unknown
        End Function

        ' no need to implement anything here if the task simply generates a SAS program
        Public ReadOnly Property RunLog() As String Implements SAS.Shared.AddIns.ISASTask.RunLog
            Get

            End Get
        End Property

        ' This property returns the SAS program that the application will run.
        Public ReadOnly Property SasCode() As String Implements SAS.Shared.AddIns.ISASTask.SasCode
            Get
                Dim builder As New StringBuilder
                builder.AppendFormat("%let data={0}.{1};" & ChrW(10), Me._consumer.ActiveData.Library, Me._consumer.ActiveData.Member)
                builder.AppendFormat("%let report={0};" & ChrW(10), Me._reportColumn)
                builder.AppendFormat("%let measure={0};" & ChrW(10), IIf((Me._statistic = eStatistic.Count), "_tpncount", Me._measureColumn))
                Dim str As String = ""
                If ((Me._statistic <> eStatistic.Count) AndAlso (Me._measureFormat.Length > 0)) Then
                    str = String.Format("%str(format={0})", Me._measureFormat)
                End If
                builder.AppendFormat("%let measureformat={0};" & ChrW(10), str)
                builder.AppendFormat("%let stat={0};" & ChrW(10), IIf((Me._statistic = eStatistic.Average), "MEAN", "SUM"))
                builder.AppendFormat("%let n={0};" & ChrW(10), Me._n.ToString(CultureInfo.InvariantCulture))
                If (Me._title.Length > 0) Then
                    builder.AppendFormat("title ""{0}"";" & ChrW(10), Me._title)
                Else
                    builder.Append("title;" & ChrW(10))
                End If
                If (Me._footnote.Length > 0) Then
                    builder.AppendFormat("footnote ""{0}"";" & ChrW(10), Me._footnote)
                Else
                    builder.Append("footnote;" & ChrW(10))
                End If
                If (Me._statistic = eStatistic.Count) Then
                    builder.AppendFormat("data work._tpnview / view=work._tpnview; " & _
                       ChrW(10) & "  set &data; " & ChrW(10) & "  _tpncount=1; " & _
                       ChrW(10) & "  label _tpncount='Count'; " & ChrW(10) & "run; " & _
                       ChrW(10) & "%let data=work._tpnview; " & ChrW(10), _
                       New Object(0 - 1) {})
                End If
                If (Me._categoryColumn.Length = 0) Then
                    builder.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopNVB.StraightReport.sas"))
                    If Me._includeChart Then
                        builder.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopNVB.StraightChart.sas"))
                    End If
                Else
                    builder.AppendFormat("%let category={0};" & ChrW(10), Me._categoryColumn)
                    builder.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopNVB.StratifiedReport.sas"))
                    If Me._includeChart Then
                        builder.Append(ReadFileFromAssembly("SASPress.CustomTasks.TopNVB.StratifiedChart.sas"))
                    End If
                End If
                Return builder.ToString
            End Get
        End Property

        Public Function Show(ByVal Owner As System.Windows.Forms.IWin32Window) As ShowResult Implements SAS.Shared.AddIns.ISASTask.Show
            Dim form As New TopNReportForm
            form.Model = Me
            form.Text = Me.Label
            If (form.ShowDialog = System.Windows.Forms.DialogResult.OK) Then
                Return 1
            End If
            Return 0
        End Function


        Public Sub Terminate() Implements SAS.Shared.AddIns.ISASTask.Terminate
            ' cleanup
        End Sub

        Private Sub ReadXML(ByVal xml As String)
            If ((Not xml Is Nothing) AndAlso (xml.Length > 0)) Then
                Try
                    Dim reader As StringReader = New StringReader(xml)
                    Dim serializer As New XmlSerializer(GetType(TopNSettings))
                    Dim settings As TopNSettings = DirectCast(serializer.Deserialize(reader), TopNSettings)
                    Me.CategoryColumn = settings.CategoryColumn
                    Me.IncludeChart = settings.IncludeChart
                    Me.MeasureColumn = settings.MeasureColumn
                    Me.N = settings.N
                    Me.ReportColumn = settings.ReportColumn
                    Me.Statistic = settings.Statistic
                    Me.Title = settings.Title
                    Me.Footnote = settings.Footnote
                    Me.MeasureFormat = settings.MeasureFormat
                Catch obj1 As System.Exception
                End Try
            End If
        End Sub

        Private Function WriteXML() As String
            Dim o As New TopNSettings
            o.CategoryColumn = Me.CategoryColumn
            o.IncludeChart = Me.IncludeChart
            o.MeasureColumn = Me.MeasureColumn
            o.N = Me.N
            o.ReportColumn = Me.ReportColumn
            o.Statistic = Me.Statistic
            o.Title = Me.Title
            o.Footnote = Me.Footnote
            o.MeasureFormat = Me.MeasureFormat
            Dim writer As StringWriter = New StringWriter
            Dim serial As New XmlSerializer(GetType(TopNSettings))
            serial.Serialize(DirectCast(writer, TextWriter), o)
            Return writer.ToString
        End Function


        ' Save and retrieve the state of this task in XML format
        Public Property XmlState() As String Implements SAS.Shared.AddIns.ISASTask.XmlState
            Get
                Return Me.WriteXML
            End Get
            Set(ByVal value As String)
                Me.ReadXML(value)
            End Set

        End Property

#End Region

#Region "ISASTaskAddIn implementation"

        Public ReadOnly Property AddInDescription() As String Implements SAS.Shared.AddIns.ISASTaskAddIn.AddInDescription
            Get
                AddInDescription = "Create a Top N report."
            End Get
        End Property

        Public ReadOnly Property AddInName() As String Implements SAS.Shared.AddIns.ISASTaskAddIn.AddInName
            Get
                AddInName = "Top N Report (VB)"
            End Get
        End Property

        Public Function Connect(ByVal Consumer As SAS.Shared.AddIns.ISASTaskConsumer) As Boolean Implements SAS.Shared.AddIns.ISASTaskAddIn.Connect
            ' Save reference to the consumer application for use later
            Me._consumer = Consumer
            Connect = True
        End Function

        Public Sub Disconnect() Implements SAS.Shared.AddIns.ISASTaskAddIn.Disconnect
            ' Free our reference to the consumer application
            Me._consumer = Nothing
        End Sub

        Public WriteOnly Property Language() As String Implements SAS.Shared.AddIns.ISASTaskAddIn.Language
            Set(ByVal Value As String)
                ' handle different languages if necessary
            End Set
        End Property

        Public Function Languages(ByRef Items() As String) As Integer Implements SAS.Shared.AddIns.ISASTaskAddIn.Languages
            ' supporting only English in this task
            Items = New String() {"en-US"}
            Languages = 1
        End Function

        Public ReadOnly Property VisibleInManager() As Boolean Implements SAS.Shared.AddIns.ISASTaskAddIn.VisibleInManager
            Get
                VisibleInManager = True
            End Get
        End Property

#End Region

#Region "ISASTaskDescription implementation"

        Public ReadOnly Property Clsid() As String Implements SAS.Shared.AddIns.ISASTaskDescription.Clsid
            Get
                ' Return GUID value generated by template
                Clsid = "103E52DE-0182-46E2-9C46-592FD5BF1EEE"
            End Get
        End Property

        Public ReadOnly Property FriendlyName() As String Implements SAS.Shared.AddIns.ISASTaskDescription.FriendlyName
            Get
                FriendlyName = "Top N Report (VB)"
            End Get
        End Property

        Public ReadOnly Property GeneratesListOutput() As Boolean Implements SAS.Shared.AddIns.ISASTaskDescription.GeneratesListOutput
            Get
                GeneratesListOutput = True
            End Get
        End Property

        Public ReadOnly Property GeneratesSasCode() As Boolean Implements SAS.Shared.AddIns.ISASTaskDescription.GeneratesSasCode
            Get
                GeneratesSasCode = True
            End Get
        End Property

        Public ReadOnly Property IconAssembly() As String Implements SAS.Shared.AddIns.ISASTaskDescription.IconAssembly
            Get
                ' return this assembly as the icon assembly
                IconAssembly = System.Reflection.Assembly.GetExecutingAssembly.Location
            End Get
        End Property

        Public ReadOnly Property IconName() As String Implements SAS.Shared.AddIns.ISASTaskDescription.IconName
            Get
                IconName = "SASPress.CustomTasks.TopNVB.CustomTask.ico"
            End Get
        End Property

        Public ReadOnly Property MajorVersion() As Integer Implements SAS.Shared.AddIns.ISASTaskDescription.MajorVersion
            Get
                MajorVersion = 1
            End Get
        End Property

        Public ReadOnly Property MinorVersion() As Integer Implements SAS.Shared.AddIns.ISASTaskDescription.MinorVersion
            Get
                MinorVersion = 0
            End Get
        End Property

        Public ReadOnly Property NumericColsRequired() As Integer Implements SAS.Shared.AddIns.ISASTaskDescription.NumericColsRequired
            Get
                NumericColsRequired = 0
            End Get
        End Property

        Public ReadOnly Property ProcsUsed() As String Implements SAS.Shared.AddIns.ISASTaskDescription.ProcsUsed
            Get
                ProcsUsed = "MEANS; REPORT; GCHART"
            End Get
        End Property

        Public ReadOnly Property ProductsOptional() As String Implements SAS.Shared.AddIns.ISASTaskDescription.ProductsOptional
            Get
                ProductsOptional = ""
            End Get
        End Property

        Public ReadOnly Property ProductsRequired() As String Implements SAS.Shared.AddIns.ISASTaskDescription.ProductsRequired
            Get
                ProductsRequired = "BASE; GRAPH"
            End Get
        End Property

        Public ReadOnly Property RequiresData() As Boolean Implements SAS.Shared.AddIns.ISASTaskDescription.RequiresData
            Get
                RequiresData = True
            End Get
        End Property

        Public ReadOnly Property StandardCategory() As Boolean Implements SAS.Shared.AddIns.ISASTaskDescription.StandardCategory
            Get
                StandardCategory = False
            End Get
        End Property

        Public ReadOnly Property TaskCategory() As String Implements SAS.Shared.AddIns.ISASTaskDescription.TaskCategory
            Get
                TaskCategory = "SAS Press Examples"
            End Get
        End Property

        Public ReadOnly Property TaskDescription() As String Implements SAS.Shared.AddIns.ISASTaskDescription.TaskDescription
            Get
                TaskDescription = "Creates a Top N report with a chart"
            End Get
        End Property

        Public ReadOnly Property TaskName() As String Implements SAS.Shared.AddIns.ISASTaskDescription.TaskName
            Get
                TaskName = "Top N Report (VB)"
            End Get
        End Property

        Public ReadOnly Property TaskType() As SAS.Shared.AddIns.ShowType Implements SAS.Shared.AddIns.ISASTaskDescription.TaskType
            Get
                TaskType = SAS.Shared.AddIns.ShowType.Wizard
            End Get
        End Property

        Public ReadOnly Property Validation() As String Implements SAS.Shared.AddIns.ISASTaskDescription.Validation
            Get
                Validation = ""
            End Get
        End Property

        Public ReadOnly Property WhatIsDescription() As String Implements SAS.Shared.AddIns.ISASTaskDescription.WhatIsDescription
            Get
                WhatIsDescription = "Creates a Top N report with a chart."
            End Get
        End Property

#End Region

        ' Constructor
        Public Sub New()
            ' initialize the label
            _label = Me.FriendlyName
        End Sub

        Friend Shared Function ReadFileFromAssembly(ByVal filename As String) As String
            Dim str As String = String.Empty
            Dim manifestResourceStream As Stream = System.Reflection.Assembly.GetCallingAssembly.GetManifestResourceStream(filename)
            If (Not manifestResourceStream Is Nothing) Then
                Dim reader As New StreamReader(manifestResourceStream)
                str = reader.ReadToEnd
                reader.Close()
                manifestResourceStream.Close()
            End If
            Return str
        End Function



    End Class

End Namespace
