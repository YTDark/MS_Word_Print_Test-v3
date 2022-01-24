Imports System.Drawing
Imports Microsoft.Office.Interop

Namespace MSO


    Public Module Print



        Public UserRequestCancelPrintingProgress As Boolean = False



        Public ReadOnly Property IsMicrosoft_Office_Insalled() As Boolean
            Get
                Dim RegKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("word.application")
                If RegKey Is Nothing Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property


        ''' <summary>
        ''' Convert a byte array to an Image
        ''' حول مصفوفة من البايتات إلى صورة
        ''' </summary>
        ''' <param name="ByteArray">مصفوفة البايتات المراد تحويلهل إلى صورة</param>
        ''' <remarks></remarks>
        Public Function ConvertBytesToImage(ByVal ByteArray() As Byte) As Image
            'Try
            If ByteArray.GetUpperBound(0) > 0 Then
                Return Image.FromStream(New System.IO.MemoryStream(ByteArray))
            Else
                Return Nothing
            End If
            'Catch ex As Exception
            '    Return Nothing
            'End Try
        End Function


    End Module


    Public Class TemplateInfo

        Public Sub New(ByVal TemplatePath As String)
            Me.m_Path = TemplatePath
        End Sub

        Private m_Path As String
        Public ReadOnly Property Path() As String
            Get
                Return Me.m_Path
            End Get
        End Property

        Private m_Document As Word.Document
        Public Property Document() As Word.Document
            Get
                Return Me.m_Document
            End Get
            Set(ByVal value As Word.Document)
                Me.m_Document = value
            End Set
        End Property

        Private m_Caption As String = ""
        Public Property Caption() As String
            Get
                Return Me.m_Caption
            End Get
            Set(ByVal value As String)
                Me.m_Caption = value
            End Set
        End Property

        Private m_PrintJob As Printing.PrintJob
        Public Property PrintJob() As Printing.PrintJob
            Get
                Return Me.m_PrintJob
            End Get
            Set(ByVal value As Printing.PrintJob)
                Me.m_PrintJob = value
            End Set
        End Property

        Private m_ViewOptions As New TemplateInfoViewOption
        Public Property ViewOptions() As MSO.TemplateInfoViewOption
            Get
                Return Me.m_ViewOptions
            End Get
            Set(ByVal value As TemplateInfoViewOption)
                Me.m_ViewOptions = value
            End Set
        End Property

    End Class

    Public Class TemplateInfoViewOption

        Public Sub New()
            With Me
                '-------------------------------------
                .ShowBookmarks = False
                .ShowTableGridlines = False
                .ArabicNumeral = MSO.Enums.MSArabicNumeral.NumeralContext
                .DisplayPageBoundaries = True
                .NormalViewDisplayRulers = True
                .ViewType = MSO.Enums.MSViewType.PrintPreview
                .WindowState = MSO.Enums.MSWindowState.Maximize
                .NormalViewZoomPageFit = MSO.Enums.MSPageFit.PageFitBestFit
                .NormalViewZoomPercentage = Nothing
                '-------------------------------------
                .PrintPreviewDisplayRulers = True
                .PrintPreviewPageFitness = New MSO.PrintPreviewPageFitness(0, 0)
                .PrintPreviewZoomPageFit = MSO.Enums.MSPageFit.PageFitBestFit
                .PrintPreviewZoomPercentage = Nothing
                '-------------------------------------
            End With
        End Sub

        Private m_ShowBookmarks As Nullable(Of Boolean)
        Public Property ShowBookmarks() As Boolean?
            Get
                Return Me.m_ShowBookmarks
            End Get
            Set(ByVal value As Boolean?)
                Me.m_ShowBookmarks = value
            End Set
        End Property

        Private m_ShowTableGridlines As Nullable(Of Boolean)
        Public Property ShowTableGridlines() As Boolean?
            Get
                Return Me.m_ShowTableGridlines
            End Get
            Set(ByVal value As Boolean?)
                Me.m_ShowTableGridlines = value
            End Set
        End Property

        Private m_DisplayPageBoundaries As Nullable(Of Boolean)
        Public Property DisplayPageBoundaries() As Boolean?
            Get
                Return Me.m_DisplayPageBoundaries
            End Get
            Set(ByVal value As Boolean?)
                Me.m_DisplayPageBoundaries = value
            End Set
        End Property

        Private m_NormalViewDisplayRulers As Nullable(Of Boolean)
        Public Property NormalViewDisplayRulers() As Boolean?
            Get
                Return Me.m_NormalViewDisplayRulers
            End Get
            Set(ByVal value As Boolean?)
                Me.m_NormalViewDisplayRulers = value
            End Set
        End Property

        Private m_PrintPreviewDisplayRulers As Nullable(Of Boolean)
        Public Property PrintPreviewDisplayRulers() As Boolean?
            Get
                Return Me.m_PrintPreviewDisplayRulers
            End Get
            Set(ByVal value As Boolean?)
                Me.m_PrintPreviewDisplayRulers = value
            End Set
        End Property

        Private m_NormalViewZoomPercentage As Nullable(Of Integer) = 100
        Public Property NormalViewZoomPercentage() As Integer?
            Get
                Return Me.m_NormalViewZoomPercentage
            End Get
            Set(ByVal value As Integer?)
                Me.m_NormalViewZoomPercentage = value
            End Set
        End Property

        Private m_PrintPreviewZoomPercentage As Nullable(Of Integer)
        Public Property PrintPreviewZoomPercentage() As Integer?
            Get
                Return Me.m_PrintPreviewZoomPercentage
            End Get
            Set(ByVal value As Integer?)
                Me.m_PrintPreviewZoomPercentage = value
            End Set
        End Property

        Private m_NormalViewZoomPageFit As Nullable(Of MSO.Enums.MSPageFit)
        Public Property NormalViewZoomPageFit() As MSO.Enums.MSPageFit?
            Get
                Return Me.m_NormalViewZoomPageFit
            End Get
            Set(ByVal value As MSO.Enums.MSPageFit?)
                Me.m_NormalViewZoomPageFit = value
            End Set
        End Property

        Private m_PrintPreviewZoomPageFit As Nullable(Of MSO.Enums.MSPageFit)
        Public Property PrintPreviewZoomPageFit() As MSO.Enums.MSPageFit?
            Get
                Return Me.m_PrintPreviewZoomPageFit
            End Get
            Set(ByVal value As MSO.Enums.MSPageFit?)
                Me.m_PrintPreviewZoomPageFit = value
            End Set
        End Property

        Private m_PrintPreviewPageFitness As PrintPreviewPageFitness
        Public Property PrintPreviewPageFitness() As PrintPreviewPageFitness
            Get
                Return Me.m_PrintPreviewPageFitness
            End Get
            Set(ByVal value As PrintPreviewPageFitness)
                Me.m_PrintPreviewPageFitness = value
            End Set
        End Property

        Private m_WindowState As Nullable(Of MSO.Enums.MSWindowState) = MSO.Enums.MSWindowState.Maximize
        Public Property WindowState() As MSO.Enums.MSWindowState?
            Get
                Return Me.m_WindowState
            End Get
            Set(ByVal value As MSO.Enums.MSWindowState?)
                Me.m_WindowState = value
            End Set
        End Property

        Private m_ViewType As Nullable(Of MSO.Enums.MSViewType)
        Public Property ViewType() As MSO.Enums.MSViewType?
            Get
                Return Me.m_ViewType
            End Get
            Set(ByVal value As MSO.Enums.MSViewType?)
                Me.m_ViewType = value
            End Set
        End Property

        Private m_ArabicNumeral As Nullable(Of MSO.Enums.MSArabicNumeral)

        Public Property ArabicNumeral() As MSO.Enums.MSArabicNumeral?
            Get
                Return Me.m_ArabicNumeral
            End Get
            Set(ByVal value As MSO.Enums.MSArabicNumeral?)
                Me.m_ArabicNumeral = value
            End Set
        End Property

    End Class

    Public Class PrintPreviewPageFitness

        Private _Rows As Integer
        Public ReadOnly Property Rows() As Integer
            Get
                Return _Rows
            End Get
        End Property

        Private _Columns As Integer
        Public ReadOnly Property Columns() As Integer
            Get
                Return _Columns
            End Get
        End Property

        Public Sub New(ByVal MaxColumns As Integer, ByVal MaxRows As Integer)
            Me._Columns = MaxColumns
            Me._Rows = MaxRows
        End Sub

    End Class



    Public Class WordApplicationNotInstalledException
        Inherits System.Exception

        Public Sub New()
            MyBase.New("جميع الكشوفات مصممة للطباعة على Microsoft Office Word " & vbNewLine &
                       "يرجى تثبيت البرنامج إذا كنت تريد طباعة الكشوفات.")
        End Sub
    End Class


End Namespace