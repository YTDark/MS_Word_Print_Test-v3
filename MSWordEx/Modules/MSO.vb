Namespace MSO

    Namespace Printing

        Public Class PrintJob

            Private m_ItemsCollection As New Collection

            Public ReadOnly Property ItemsCollection() As Collection
                Get
                    Return Me.m_ItemsCollection
                End Get
            End Property

            Public Sub AddText(ByVal Text As String, ByVal BookMarkName As String)
                'Try
                If Not String.IsNullOrEmpty(Text) AndAlso Not String.IsNullOrEmpty(BookMarkName) Then
                    Dim PrintItem As New Printing.Text
                    PrintItem.BookMarkName = BookMarkName
                    PrintItem.Text = Text
                    Me.ItemsCollection.Add(PrintItem)
                End If

                'Catch ex As Exception
                '    MSO.PrintingProcess.ShowErrorMsgAndClose(ex.Message)
                'End Try
            End Sub

            Public Function AddImage() As Printing.Image
                'Try
                Dim PrintItem As New Printing.Image
                Me.ItemsCollection.Add(PrintItem)
                Return PrintItem
                'Catch ex As Exception
                '    MSO.PrintingProcess.ShowErrorMsgAndClose(ex.Message)
                '    Return Nothing
                'End Try
            End Function

            Public Function AddTable() As Printing.Table
                'Try
                Dim PrintItem As New Printing.Table
                    Me.ItemsCollection.Add(PrintItem)
                    Return PrintItem
                'Catch ex As Exception
                '    MSO.PrintingProcess.ShowErrorMsgAndClose(ex.Message)
                '    Return Nothing
                'End Try
            End Function

            Public Function GetPrintJobCuonter() As Integer
                'Try
                Dim PrintJobCuonter As Integer = 0
                For Each PrintItem As Object In Me.ItemsCollection
                    If TypeOf PrintItem Is Printing.Table Then
                        Dim Table As Printing.Table = PrintItem

                        Dim AutoNumberCount As Integer = 0
                        If Table.IsFirstColumnAutoNumber Then AutoNumberCount = 1
                        PrintJobCuonter += ((Table.Columns.Count + AutoNumberCount) * Table.DataTable.Rows.Count)
                    ElseIf TypeOf PrintItem Is Printing.Text OrElse TypeOf PrintItem Is Printing.Image Then
                        PrintJobCuonter += 1
                    End If
                Next
                Return PrintJobCuonter
                'Catch ex As Exception
                '    MSO.PrintingProcess.ShowErrorMsgAndClose(ex.Message)
                'End Try
            End Function

        End Class

        Public Class Text
            Private m_BookMarkName As String
            Public Property BookMarkName() As String
                Get
                    Return Me.m_BookMarkName
                End Get
                Set(ByVal value As String)
                    Me.m_BookMarkName = value
                End Set
            End Property
            Private m_Text As String
            Public Property Text() As String
                Get
                    Return Me.m_Text
                End Get
                Set(ByVal value As String)
                    Me.m_Text = value
                End Set
            End Property
        End Class

        Public Class Image

            Private m_BookMarkName As String
            Public Property BookMarkName() As String
                Get
                    Return Me.m_BookMarkName
                End Get
                Set(ByVal value As String)
                    Me.m_BookMarkName = value
                End Set
            End Property


            Private m_Image As System.Drawing.Image
            Public Property Image() As System.Drawing.Image
                Get
                    Return Me.m_Image
                End Get
                Set(ByVal value As System.Drawing.Image)
                    Me.m_Image = value
                End Set
            End Property


            Private m_MaximumSize As System.Drawing.Size
            Public Property MaximumSize() As System.Drawing.Size
                Get
                    Return Me.m_MaximumSize
                End Get
                Set(ByVal value As System.Drawing.Size)
                    Me.m_MaximumSize = value
                End Set
            End Property


            Private m_BorderStyle As MSO.Enums.MSLineStyle
            Public Property BorderStyle() As MSO.Enums.MSLineStyle
                Get
                    Return Me.m_BorderStyle
                End Get
                Set(ByVal value As MSO.Enums.MSLineStyle)
                    Me.m_BorderStyle = value
                End Set
            End Property


            Private m_BorderSize As MSO.Enums.MSLineWidth
            Public Property BorderSize() As MSO.Enums.MSLineWidth
                Get
                    Return Me.m_BorderSize
                End Get
                Set(ByVal value As MSO.Enums.MSLineWidth)
                    Me.m_BorderSize = value
                End Set
            End Property


            Private m_BorderColor As MSO.Enums.MSColor = MSO.Enums.MSColor.Black
            Public Property BorderColor() As MSO.Enums.MSColor
                Get
                    Return Me.m_BorderColor
                End Get
                Set(ByVal value As MSO.Enums.MSColor)
                    Me.m_BorderColor = value
                End Set
            End Property

        End Class

        Public Class Table

            Private m_DataTable As System.Data.DataTable
            Public Property DataTable() As System.Data.DataTable
                Get
                    Return Me.m_DataTable
                End Get
                Set(ByVal value As System.Data.DataTable)
                    Me.m_DataTable = value
                End Set
            End Property

            Private m_TableHeadBookMarkName As String = ""
            Public Property TableHeadBookMarkName() As String
                Get
                    Return Me.m_TableHeadBookMarkName
                End Get
                Set(ByVal value As String)
                    Me.m_TableHeadBookMarkName = value
                End Set
            End Property

            Private m_FirstRowBookMarkName As String = ""
            Public Property FirstRowBookMarkName() As String
                Get
                    Return Me.m_FirstRowBookMarkName
                End Get
                Set(ByVal value As String)
                    Me.m_FirstRowBookMarkName = value
                End Set
            End Property


            Private m_IsFirstColumnAutoNumber As Boolean = False
            Public Property IsFirstColumnAutoNumber() As Boolean
                Get
                    Return Me.m_IsFirstColumnAutoNumber
                End Get
                Set(ByVal value As Boolean)
                    Me.m_IsFirstColumnAutoNumber = value
                End Set
            End Property


            Private m_MinimumRowsAtTheBeginningOfTable As Integer = 1 ' MinimumNumberOfRowsUnderTheTableHeadInSinglePage
            Public Property MinimumRowsAtTheBeginningOfTable() As Integer
                Get
                    Return m_MinimumRowsAtTheBeginningOfTable
                End Get
                Set(ByVal value As Integer)
                    m_MinimumRowsAtTheBeginningOfTable = value
                End Set
            End Property


            Private m_DeleteTableIfNoData As Boolean = False
            Public Property DeleteTableIfNoData() As Boolean
                Get
                    Return Me.m_DeleteTableIfNoData
                End Get
                Set(ByVal value As Boolean)
                    Me.m_DeleteTableIfNoData = value
                End Set
            End Property



            Private m_Columns As New Microsoft.VisualBasic.Collection
            Public ReadOnly Property Columns() As Microsoft.VisualBasic.Collection
                Get
                    Return Me.m_Columns
                End Get
            End Property



            '##################################################
            Public Function AddEmptyColumn() As ColumnsItem.EmptyColumn
                Dim NewEmptyColumn As New ColumnsItem.EmptyColumn
                Me.Columns.Add(NewEmptyColumn)
                Return NewEmptyColumn
            End Function
            '##################################################
            Public Function AddImageColumn(ByVal ColumnName As String) As ColumnsItem.ImageColumn
                Dim NewImageColumn As New ColumnsItem.ImageColumn(ColumnName)
                Me.Columns.Add(NewImageColumn)
                Return NewImageColumn
            End Function
            '##################################################
            Public Function AddTextColumn(ByVal ColumnName As String) As ColumnsItem.TextColumn
                Dim NewTextColumn As New ColumnsItem.TextColumn(ColumnName)
                Me.Columns.Add(NewTextColumn)
                Return NewTextColumn
            End Function


        End Class

        Public Class ImageStyle
            Public MaximumSize As System.Drawing.Size
            Public BorderStyle As MSO.Enums.MSLineStyle
            Public BorderColor As MSO.Enums.MSColor
            Public BorderSize As MSO.Enums.MSLineWidth
        End Class

        Namespace ColumnsItem

            Public Class TextColumn
                Private m_Name As String
                Public ReadOnly Property Name() As String
                    Get
                        Return Me.m_Name
                    End Get
                End Property
                Public Sub New(ByVal ColumnName As String)
                    Me.m_Name = ColumnName
                End Sub
            End Class

            Public Class ImageColumn

                Private m_Name As String
                Public ReadOnly Property Name() As String
                    Get
                        Return Me.m_Name
                    End Get
                End Property

                Private m_SourceType As MSO.Enums.ImageSource
                Public Property SourceType() As MSO.Enums.ImageSource
                    Get
                        Return Me.m_SourceType
                    End Get
                    Set(ByVal value As MSO.Enums.ImageSource)
                        Me.m_SourceType = value
                    End Set
                End Property

                Private m_Style As New ImageStyle
                Public Property Style() As ImageStyle
                    Get
                        Return Me.m_Style
                    End Get
                    Set(ByVal value As ImageStyle)
                        Me.m_Style = value
                    End Set
                End Property

                Public Sub New(ByVal ColumnName As String)
                    Me.m_Name = ColumnName
                End Sub

            End Class

            Public Class EmptyColumn
            End Class

        End Namespace

    End Namespace

End Namespace





