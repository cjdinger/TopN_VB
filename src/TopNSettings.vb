Imports SASPress.CustomTasks.TopNVB

<Serializable()> _
Public Class TopNSettings
    ' Properties
    Public Property CategoryColumn() As String
        Get
            Return Me._categoryColumn
        End Get
        Set(ByVal value As String)
            Me._categoryColumn = value
        End Set
    End Property

    Public Property Footnote() As String
        Get
            Return Me._footnote
        End Get
        Set(ByVal value As String)
            Me._footnote = value
        End Set
    End Property

    Public Property IncludeChart() As Boolean
        Get
            Return Me._includeChart
        End Get
        Set(ByVal value As Boolean)
            Me._includeChart = value
        End Set
    End Property

    Public Property MeasureColumn() As String
        Get
            Return Me._measureColumn
        End Get
        Set(ByVal value As String)
            Me._measureColumn = value
        End Set
    End Property

    Public Property MeasureFormat() As String
        Get
            Return Me._measureFormat
        End Get
        Set(ByVal value As String)
            Me._measureFormat = value
        End Set
    End Property

    Public Property N() As Integer
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

    Public Property ReportColumn() As String
        Get
            Return Me._reportColumn
        End Get
        Set(ByVal value As String)
            Me._reportColumn = value
        End Set
    End Property

    Public Property Statistic() As SASPress.CustomTasks.TopNVB.TopNReport.eStatistic
        Get
            Return Me._statistic
        End Get
        Set(ByVal value As SASPress.CustomTasks.TopNVB.TopNReport.eStatistic)
            Me._statistic = value
        End Set
    End Property

    Public Property Title() As String
        Get
            Return Me._title
        End Get
        Set(ByVal value As String)
            Me._title = value
        End Set
    End Property


    ' Fields
    Private _categoryColumn As String = ""
    Private _footnote As String = ""
    Private _includeChart As Boolean = True
    Private _measureColumn As String = ""
    Private _measureFormat As String = ""
    Private _n As Integer = 10
    Private _reportColumn As String = ""
    Private _statistic As SASPress.CustomTasks.TopNVB.TopNReport.eStatistic = SASPress.CustomTasks.TopNVB.TopNReport.eStatistic.Sum
    Private _title As String = ""
End Class



