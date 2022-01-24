

Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Namespace MSO

    Public Class MSWord

        Public WithEvents WordApplication As Word.Application
        Public TemplatesInfoCollection As New Collections.ObjectModel.Collection(Of TemplateInfo)
        Private WithEvents CurrentTempInfo As TemplateInfo
        Public WithEvents ParentForm As System.Windows.Forms.Form
        Private WithEvents ProgressDialog As PrintingProgressDialog
        Private ImagePath As String = String.Empty
        Public Sub New(ByRef ParentForm As System.Windows.Forms.Form)


            If Not MSO.IsMicrosoft_Office_Insalled Then
                Throw New WordApplicationNotInstalledException()
            End If



            Me.ParentForm = ParentForm

            Me.WordApplication = New Word.Application
            Me.WordApplication.Visible = False
            'Me.ParentForm = ParentForm
            Me.ImagePath = My.Computer.FileSystem.SpecialDirectories.Temp & "\Image"

        End Sub


        Public Function AppIsDead() As Boolean

            Try
                Dim DocCount As Integer = Me.WordApplication.Documents.Count
                Return False
            Catch ex As Exception
                Return True
            End Try

        End Function


        Public Sub TryShow()
            Me.WordApplication.Activate()
            Me.WordApplication.ShowMe()
            Me.WordApplication.ShowWindowsInTaskbar = True
            Me.WordApplication.Visible = True
            Me.WordApplication.WindowState = WdWindowState.wdWindowStateMaximize
        End Sub

        Public Sub ClearAllDocuments()

            For I As Integer = 1 To Me.WordApplication.Documents.Count
                Try
                    Me.WordApplication.Documents(I).Close(False)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Clear Documents...")
                End Try
            Next

        End Sub



        '-- هذا الأمر فقط لتشغيل الوورد بمستند فاضي وعمل إجراءات
        '-- السبب هو عندما يحتاج المستخدم لمثل هذا الإجراء لن ينتظر طويلاً, لأنه تم عمل مرة واحدة من قبل عند أول تشغيل للبرنامج
        '-------------------------------------------------------------------------------------
        Public Sub RunEmptyDocAndExportAsPdf(DocPath As String, PdfPath As String)
            Me.WordApplication.Documents.Add(DocPath, False)
            Me.ExportAsPdfFormat(PdfPath)
        End Sub
        '-------------------------------------------------------------------------------------


        Public Function AddNewTemplateInfo(ByVal TemplateInfo As TemplateInfo) As Boolean
            'Try

            TemplateInfo.Document = Me.WordApplication.Documents.Add(TemplateInfo.Path, False)
            Me.TemplatesInfoCollection.Add(TemplateInfo)

            Return True
            'Catch ex As Exception
            '    MSO.PrintingProcess.ShowErrorMsgAndClose(ex.Message)
            '    Return False
            'End Try

        End Function

        Public Sub CloseApplication()
            Me.ClearAllDocuments()
            '----------------------------
            Try
                Me.WordApplication.Quit(False)
                Me.WordApplication = Nothing
            Catch ex As Exception
            End Try
        End Sub

        Public Sub SaveAs(ByVal Output_Filename As String)
            Me.WordApplication.ActiveDocument.SaveAs2(Output_Filename)
        End Sub


        Public Sub ExportAsPdfFormat(ByVal Output_Filename As String)
            If Not Me.WordApplication.ActiveDocument Is Nothing Then
                Me.WordApplication.ActiveDocument.ExportAsFixedFormat(Output_Filename,
                                                                    WdExportFormat.wdExportFormatPDF, False,
                                                                    WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                                                                    WdExportRange.wdExportAllDocument, 0, 0,
                                                                    WdExportItem.wdExportDocumentContent, False, False,
                                                                    WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)
            End If
        End Sub



        Public Function PrintOut(Optional ByVal PrinterName As String = "") As Boolean
            Try
                If PrinterName.Trim = "" Then
                    Me.WordApplication.PrintOut()
                Else
                    Dim PrinterExists As Boolean = False
                    For i = 0 To Drawing.Printing.PrinterSettings.InstalledPrinters.Count - 1
                        If PrinterName.ToLower = Drawing.Printing.PrinterSettings.InstalledPrinters.Item(i).ToLower Then
                            PrinterExists = True
                            Exit For
                        End If
                    Next
                    If PrinterExists Then
                        Me.WordApplication.WordBasic.FilePrintSetup(Printer:=PrinterName, DoNotSetAsSysDefault:=1)
                        Me.WordApplication.PrintOut()
                    Else
                        MsgBox("The name of the printer does not exist.", MsgBoxStyle.Exclamation)
                    End If
                End If
                Using WaitThePrinter As New Frm_WaitThePrinterToReceiveTheFile
                    WaitThePrinter.WaitThePrinterToReceiveTheFile(Me.WordApplication)
                End Using

                Return True
            Catch ex As Exception
                Throw New Exception(ex.Message)
                Return False
            End Try
        End Function

        Public Function PrintShowDialog() As Boolean
            Dim p As New PrintDialog
            If p.ShowDialog = DialogResult.OK Then
                Me.WordApplication.WordBasic.FilePrintSetup(Printer:=p.PrinterSettings.PrinterName, DoNotSetAsSysDefault:=1)
                Me.WordApplication.PrintOut()

                Using WaitThePrinter As New Frm_WaitThePrinterToReceiveTheFile
                    WaitThePrinter.WaitThePrinterToReceiveTheFile(Me.WordApplication)
                End Using
                Return True
            Else
                Return False
            End If
        End Function

        Public Overloads Sub Dispose()
            Me.Finalize()
            GC.SuppressFinalize(Me)
        End Sub

        Public Sub PrintPreview()
            Me.PrepairTemplatesCollection(True)
        End Sub






        Public Sub Quit(Optional ByVal SaveChanges As Boolean = False)
            Me.WordApplication.Quit(SaveChanges)
        End Sub

        Public Sub RunReportForTestint()
            Me.WordApplication.ScreenUpdating = False

            For Each TempInfo As TemplateInfo In Me.TemplatesInfoCollection
                Me.WordApplication.Documents.Item(TempInfo.Document).Activate()
                Me.CurrentTempInfo = TempInfo
                If TempInfo.Caption.Trim <> "" Then
                    Me.CurrentTempInfo.Document.Application.ActiveWindow.Caption = TempInfo.Caption
                End If
                If Not DoRepairingJob(TempInfo) Then
                    Continue For
                End If
                SetViewOptionForThisDocumentInfo(TempInfo)
                TempInfo.Document.Saved = True
                System.Windows.Forms.Application.DoEvents()
            Next
            Me.WordApplication.ScreenUpdating = True
            System.Windows.Forms.Application.DoEvents()
            Try
                IO.File.Delete(Me.ImagePath)
            Catch ex As Exception
            End Try
        End Sub

        Public Function PrepairTemplatesCollection(Optional ByVal ShowMSWordAfterFinish As Boolean = True) As Boolean

            MSO.Print.UserRequestCancelPrintingProgress = False
            Me.ProgressDialog = New PrintingProgressDialog()
            Me.ProgressDialog.ProgressBarEx1.Maximum = GetTotalJobCounter()

            If Me.ProgressDialog.ShowDialog(Me.ParentForm) <> DialogResult.OK Then
                Return False
            End If

            If ShowMSWordAfterFinish Then
                Me.WordApplication.Visible = True
            End If

            Return True

        End Function

        Public Sub SetMSWordVisible(ByVal Visible As Boolean)
            Me.WordApplication.Visible = Visible
        End Sub


        Private Function GetTotalJobCounter() As Integer
            Dim Total As Integer = 0
            For Each TempInfo As TemplateInfo In Me.TemplatesInfoCollection
                Total += TempInfo.PrintJob.GetPrintJobCuonter
            Next
            If Total <= 0 Then
                Return 0
            End If

            Return Total

        End Function


        Private Sub ProgressDialog_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProgressDialog.Shown

            If Not DoRepairTemplatesCollection() Then
                Me.ProgressDialog.DialogResult = DialogResult.Abort
                Me.ProgressDialog.Close()
            End If

        End Sub


        Private Function DoRepairTemplatesCollection() As Boolean

            Me.WordApplication.ScreenUpdating = False

            For Each TempInfo As TemplateInfo In Me.TemplatesInfoCollection

                Me.ProgressDialog.Tbox_ActionName.Text = TempInfo.Caption

                Me.WordApplication.Documents.Item(TempInfo.Document).Activate()

                Me.CurrentTempInfo = TempInfo

                If TempInfo.Caption.Trim <> "" Then
                    Me.CurrentTempInfo.Document.Application.ActiveWindow.Caption = TempInfo.Caption
                End If

                '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

                If Not DoRepairingJob(TempInfo) Then
                    Return False
                End If

                '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

                SetViewOptionForThisDocumentInfo(TempInfo)

                TempInfo.Document.Saved = True

                System.Windows.Forms.Application.DoEvents()

            Next

            Me.ProgressDialog.Hide()
            Me.ProgressDialog.DialogResult = DialogResult.OK
            Me.ProgressDialog.Close()
            Me.WordApplication.ScreenUpdating = True

            System.Windows.Forms.Application.DoEvents()

            Try
                IO.File.Delete(Me.ImagePath)
            Catch ex As Exception
            End Try

            Return True

        End Function


        Private Function DoRepairingJob(ByVal TempInfo As TemplateInfo) As Boolean

            '-----------------------------------------------
            For Each Item As Object In TempInfo.PrintJob.ItemsCollection

                If MSO.Print.UserRequestCancelPrintingProgress Then
                    Return False
                End If

                '------------------------------------------------------------------------------------
                If TypeOf Item Is Printing.Text Then
                    Dim PrintItem As Printing.Text = Item
                    If Not SelectThisBookMark(PrintItem.BookMarkName) Then Return False
                    TempInfo.Document.Application.ActiveWindow.Selection.TypeText(Text:=PrintItem.Text)
                    SetProgressCuonter()
                    System.Windows.Forms.Application.DoEvents()
                    '--------------------------------------------------------------------------------
                ElseIf TypeOf Item Is Printing.Image Then
                    Dim PrintItem As Printing.Image = Item
                    If Not String.IsNullOrEmpty(PrintItem.BookMarkName) AndAlso PrintItem.Image IsNot Nothing Then
                        If Not SelectThisBookMark(PrintItem.BookMarkName) Then Return False
                    End If
                    WriteImage(PrintItem)
                    SetProgressCuonter()
                    System.Windows.Forms.Application.DoEvents()
                    '--------------------------------------------------------------------------------
                ElseIf TypeOf Item Is Printing.Table Then
                    Dim PrintItem As Printing.Table = Item
                    If PrintItem.DataTable IsNot Nothing Then
                        If PrintItem.DataTable.Rows.Count <= 0 Then
                            If PrintItem.DeleteTableIfNoData = True Then
                                If Not SelectThisBookMark(PrintItem.TableHeadBookMarkName) Then Return False
                                Me.WordApplication.Selection.Tables(1).Select()
                                Me.WordApplication.Selection.Tables(1).Delete()
                                System.Windows.Forms.Application.DoEvents()
                            End If
                        Else
                            If Not WriteTableContaint(PrintItem) Then Return False
                        End If
                    End If
                End If
                '------------------------------------------------------------------------------------
            Next

            TempInfo.Document.Application.ActiveWindow.Selection.HomeKey(Unit:=Word.WdUnits.wdStory)
            TempInfo.Document.Application.ActiveWindow.Selection.MoveUp(Unit:=Word.WdUnits.wdScreen, Count:=1)
            System.Windows.Forms.Application.DoEvents()
            '-----------------------------------------------
            Return True

        End Function


        Private Sub WriteImage(ByVal PrintImage As Printing.Image)
            '-- 
            Dim scale_factor As Single = 1.0
            Dim ImageHasReSized As Boolean = False

            If Not PrintImage.MaximumSize.IsEmpty Then
                '-------------------------------------------------------------------------------
                If PrintImage.Image.Height > PrintImage.MaximumSize.Height Then
                    ' Reduce height first
                    scale_factor = CSng(PrintImage.MaximumSize.Height / PrintImage.Image.Height)
                    ImageHasReSized = True
                End If
                '---------------------------------------------------------------------------------
                If (PrintImage.Image.Width * scale_factor) > PrintImage.MaximumSize.Width Then
                    ' Scaled width exceeds Box's width, recalculate scale_factor
                    scale_factor = CSng(PrintImage.MaximumSize.Width / PrintImage.Image.Width)
                    ImageHasReSized = True
                End If
                '-------------------------------------------------------------------------------
            End If

            If ImageHasReSized Then
                '-------------------------------------------------------------------------------
                Dim ImageAfterResizing As New Bitmap(PrintImage.Image,
                                                CInt(PrintImage.Image.Width * scale_factor),
                                                CInt(PrintImage.Image.Height * scale_factor))
                ImageAfterResizing.Save(Me.ImagePath, System.Drawing.Imaging.ImageFormat.Jpeg)
            Else
                PrintImage.Image.Save(Me.ImagePath, System.Drawing.Imaging.ImageFormat.Jpeg)
            End If

            Me.WordApplication.ActiveWindow.Selection.InlineShapes.AddPicture(Me.ImagePath)

            If PrintImage.BorderStyle <> Word.WdLineStyle.wdLineStyleNone Then
                SetImageBorder(PrintImage.BorderStyle, PrintImage.BorderSize, PrintImage.BorderColor)
            End If
        End Sub


        Private Sub SetImageBorder(ByVal BorderStyle As Word.WdLineStyle, ByVal BorderSize As Word.WdLineWidth, ByVal BorderColor As Word.WdColor)
            'Try
            Dim PositevBorder As Word.WdLineStyle = BorderStyle
            Dim NegitevBorder As Word.WdLineStyle = GetNigitevBorder(BorderStyle)
            With Me.WordApplication.ActiveDocument.InlineShapes(Me.WordApplication.ActiveDocument.InlineShapes.Count)
                With .Borders(Word.WdBorderType.wdBorderTop)
                    .LineStyle = PositevBorder
                    .LineWidth = BorderSize
                    .Color = BorderColor
                End With
                With .Borders(Word.WdBorderType.wdBorderLeft)
                    .LineStyle = PositevBorder
                    .LineWidth = BorderSize
                    .Color = BorderColor
                End With
                With .Borders(Word.WdBorderType.wdBorderRight)
                    .LineStyle = NegitevBorder
                    .LineWidth = BorderSize
                    .Color = BorderColor
                End With
                With .Borders(Word.WdBorderType.wdBorderBottom)
                    .LineStyle = NegitevBorder
                    .LineWidth = BorderSize
                    .Color = BorderColor
                End With
            End With
            'Catch ex As Exception
            'End Try
        End Sub


        Private Function GetNigitevBorder(ByVal BorderStyle As Word.WdLineStyle) As Word.WdLineStyle
            Select Case BorderStyle
                '--------------------------------------------------------
                Case Is = Word.WdLineStyle.wdLineStyleThinThickSmallGap
                    Return Word.WdLineStyle.wdLineStyleThickThinSmallGap
                Case Is = Word.WdLineStyle.wdLineStyleThickThinSmallGap
                    Return Word.WdLineStyle.wdLineStyleThinThickSmallGap
                    '----------------------------------------------------
                Case Is = Word.WdLineStyle.wdLineStyleThinThickMedGap
                    Return Word.WdLineStyle.wdLineStyleThickThinMedGap
                Case Is = Word.WdLineStyle.wdLineStyleThickThinMedGap
                    Return Word.WdLineStyle.wdLineStyleThinThickMedGap
                    '----------------------------------------------------
                Case Is = Word.WdLineStyle.wdLineStyleEmboss3D
                    Return Word.WdLineStyle.wdLineStyleEngrave3D
                Case Is = Word.WdLineStyle.wdLineStyleEngrave3D
                    Return Word.WdLineStyle.wdLineStyleEmboss3D
                    '----------------------------------------------------
                Case Is = Word.WdLineStyle.wdLineStyleThinThickLargeGap
                    Return Word.WdLineStyle.wdLineStyleThickThinLargeGap
                Case Is = Word.WdLineStyle.wdLineStyleThickThinLargeGap
                    Return Word.WdLineStyle.wdLineStyleThinThickLargeGap
                    '----------------------------------------------------
                Case Else
                    Return BorderStyle
            End Select
        End Function


        Private Sub SetProgressCuonter()

            Me.ProgressDialog.SetProgressCuonter()
            'Me.ProgressDialog.PrintProgressBar.Value += 1
            'Dim Percentage As Double = Me.ProgressDialog.PrintProgressBar.Value / Me.ProgressDialog.PrintProgressBar.Maximum * 100
            'Me.ProgressDialog.lbl_CounterPercentage.Text = Math.Round(Percentage, 0).ToString & "%"
            'System.Windows.Forms.Application.DoEvents()
        End Sub


        Private Function SelectThisBookMark(ByVal BookMarkName As String) As Boolean

            With Me.WordApplication.ActiveDocument.Bookmarks
                If .Exists(BookMarkName) = True Then
                    .Item(BookMarkName).Select()
                    Return True
                Else

                    Throw New Exception("MSO Document BookMark Name [ " & BookMarkName & "] Not Exists!")
                    Return False
                End If
            End With

        End Function


        Private RndCount As Integer = 0
        Private Function GetRandomBookMarkName() As String
            RndCount += 1
            Return "BM_Rnd_" & RndCount
        End Function


        Private Function WriteTableContaint(ByVal PrintItem As Printing.Table) As Boolean

            Dim Cuonter As Integer = 0                        '<--- عداد الترقيم التلقائي
            Dim CurrentPage As Integer = 0                    '<--- رقم الصفحة الحالية
            Dim StartBlockLocation As Integer = 0             '<--- رقم الصفحة عند عنوان الجدول الرئيسي
            Dim TableStartPointLocation As Integer = 0        '<--- رقم الصفحة عند أول سطر في الجول
            Dim BookMarkNameOfMinimumNumberOfRows As String = ""
            'Try


            '>>>> Start Writing Data Into the Table >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '-- كتابة البيانات في الجدول
            If Not SelectThisBookMark(PrintItem.TableHeadBookMarkName) Then Return False
            StartBlockLocation = Me.WordApplication.ActiveWindow.Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber)

            For Each dr As Data.DataRow In PrintItem.DataTable.Rows


                If MSO.Print.UserRequestCancelPrintingProgress Then
                    Return False
                End If

                Cuonter += 1
                If Cuonter = 1 Then
                    If Not SelectThisBookMark(PrintItem.FirstRowBookMarkName) Then Return False
                    TableStartPointLocation = Me.WordApplication.ActiveWindow.Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber)
                Else
                    Me.WordApplication.ActiveWindow.Selection.MoveRight(Unit:=Word.WdUnits.wdCell)
                End If

                If PrintItem.MinimumRowsAtTheBeginningOfTable > 1 Then
                    If Cuonter = PrintItem.MinimumRowsAtTheBeginningOfTable Then

                        BookMarkNameOfMinimumNumberOfRows = Me.GetRandomBookMarkName()
                        With Me.WordApplication.ActiveDocument
                            .Bookmarks.Add(
                                            Range:=Me.WordApplication.ActiveWindow.Selection.Range,
                                            Name:=BookMarkNameOfMinimumNumberOfRows)
                        End With
                    End If
                End If


                If PrintItem.IsFirstColumnAutoNumber Then
                    Me.WordApplication.ActiveWindow.Selection.TypeText(Text:=Cuonter & ".")
                    Me.WordApplication.ActiveWindow.Selection.MoveRight(Unit:=Word.WdUnits.wdCell)
                    SetProgressCuonter()
                End If


                Dim ColumnCuonter As Integer = 0
                '########################################################################################
                For Each ThisColumn As Object In PrintItem.Columns

                    ColumnCuonter += 1

                    '    ---------------------------------

                    If TypeOf ThisColumn Is MSO.Printing.ColumnsItem.TextColumn Then

                        Dim Column As MSO.Printing.ColumnsItem.TextColumn = DirectCast(ThisColumn, MSO.Printing.ColumnsItem.TextColumn)

                        Me.WordApplication.ActiveWindow.Selection.TypeText(Text:=dr(Column.Name).ToString)

                        '----------------------------------

                    ElseIf TypeOf ThisColumn Is MSO.Printing.ColumnsItem.ImageColumn Then

                        Try
                            Dim Column As MSO.Printing.ColumnsItem.ImageColumn = DirectCast(ThisColumn, MSO.Printing.ColumnsItem.ImageColumn)

                            Dim PrintImage As New Printing.Image

                            If Column.SourceType = MSO.Enums.ImageSource.BytesArray Then

                                PrintImage.Image = ConvertBytesToImage(dr(Column.Name))

                            ElseIf Column.SourceType = MSO.Enums.ImageSource.FilePath Then

                                PrintImage.Image = Image.FromFile(dr(Column.Name).ToString)

                            End If

                            PrintImage.BorderColor = Column.Style.BorderColor
                            PrintImage.BorderSize = Column.Style.BorderSize
                            PrintImage.BorderStyle = Column.Style.BorderStyle
                            PrintImage.MaximumSize = Column.Style.MaximumSize

                            WriteImage(PrintImage)
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try


                        '----------------------------------

                    ElseIf TypeOf ThisColumn Is MSO.Printing.ColumnsItem.EmptyColumn Then
                        ' Do Nothing , Gust Go To the Next One
                    End If

                    If Not (ColumnCuonter = PrintItem.Columns.Count) Then
                        Me.WordApplication.ActiveWindow.Selection.MoveRight(Unit:=Word.WdUnits.wdCell)
                    End If
                    SetProgressCuonter()
                    System.Windows.Forms.Application.DoEvents()
                Next
                '########################################################################################

                System.Windows.Forms.Application.DoEvents()
            Next
            '<<<< End Writing Data Into the Table <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


            '>>>> Start Repair Table Head Location >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '-- هذا القسم يختص بمعالجة مكان رأس الجدول
            '-- إذا كان ( مستند الوورد - أو القالب الذي صممتة والذي يتم كتابة البيانات فية الآن ) يحتوي على أكثر من جدول
            '-- لنفترض ان الجدول الأول وبعد الإنتهاء من الكتابة فية وصل إلى نهاية الصفحة من الأسفل  
            '-- وتبقى مكان صغير في أسفل الصفحة من الممكن ان يبداء رأس الجدول الثاني من هذا المكان الصغير
            '-- في هذة الحالة من الممكن ان يكون رأس الجدول الثاني في نهاية الصفحة ثم يتم كتابة البيانات في الصفحة التالية
            '-- في هذة الحالة ستكون الصفحة التالية لا تحتوي على رأس الجدول
            '-- في هذا القسم يتم تحريك الجدول إلى الصفحة التالية لتبداء برأس الجدول ثم البيانات مباشرتاً
            '--------------------------------------------------------------------
            If Cuonter >= PrintItem.MinimumRowsAtTheBeginningOfTable AndAlso PrintItem.MinimumRowsAtTheBeginningOfTable > 1 Then
                If Not SelectThisBookMark(BookMarkNameOfMinimumNumberOfRows) Then Return False
            Else
                If Not SelectThisBookMark(PrintItem.FirstRowBookMarkName) Then Return False
            End If
            Dim MinimumNumberOfRows_Location As Integer = Me.WordApplication.ActiveWindow.Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber)
            '--------------------------------------------------------------------
            If Not SelectThisBookMark(PrintItem.TableHeadBookMarkName) Then Return False
            StartBlockLocation = Me.WordApplication.ActiveWindow.Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber)
            '--------------------------------------------------------------------
            If StartBlockLocation <> MinimumNumberOfRows_Location Then
                Dim NewLocation As Integer = StartBlockLocation + 1
                Do
                    If Not SelectThisBookMark(PrintItem.TableHeadBookMarkName) Then Return False
                    Me.WordApplication.ActiveWindow.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=1)
                    Me.WordApplication.ActiveWindow.Selection.TypeParagraph()
                    If Not SelectThisBookMark(PrintItem.TableHeadBookMarkName) Then Return False
                    StartBlockLocation = Me.WordApplication.ActiveWindow.Selection.Information(Word.WdInformation.wdActiveEndAdjustedPageNumber)

                    System.Windows.Forms.Application.DoEvents()

                Loop Until StartBlockLocation = NewLocation
            End If
            '<<<< End Repair Table Head Location <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            Return True
            'Catch ex As Exception
            '    Return False
            'End Try

        End Function


        Private Sub Application_WindowActivate(ByVal Doc As Word.Document, ByVal Wn As Word.Window) Handles WordApplication.WindowActivate
            If Me.WordApplication.ActiveWindow.View.Type = Word.WdViewType.wdPrintPreview Then
                Me.WordApplication.ActiveWindow.View.Magnifier = True
            End If
        End Sub

        Private Sub SetPrintPreviewPageFitness(ByVal MaxColumns As Integer, ByVal MaxRows As Integer)

            '---------------------------------------------------
            If MaxColumns <= 0 OrElse MaxRows <= 0 Then Exit Sub
            '---------------------------------------------------

            Dim PageCount, Remainder, Columns, Rows As Integer

            PageCount = Me.WordApplication.ActiveDocument.ComputeStatistics(Word.WdStatistic.wdStatisticPages)

            Rows = Math.DivRem(PageCount, MaxColumns, Remainder)

            If PageCount > MaxColumns Then
                Columns = MaxColumns
                If Remainder <> 0 Then
                    Rows += 1
                    If Rows > MaxRows Then
                        Rows = MaxRows
                    End If
                End If
            Else
                Columns = PageCount
                Rows = 1
            End If

            Me.WordApplication.ActiveWindow.View.Zoom.PageColumns = Columns
            Me.WordApplication.ActiveWindow.View.Zoom.PageRows = Rows

        End Sub


        Private Sub SetViewOptionForThisDocumentInfo(ByVal TemplateInfo As TemplateInfo)

            With TemplateInfo.ViewOptions

                '------------------------------------------------------------------------------------------
                If .WindowState.HasValue Then
                    Me.WordApplication.ActiveWindow.WindowState = .WindowState.Value
                End If
                '------------------------------------------------------------------------------------------
                If .ArabicNumeral.HasValue Then
                    Me.WordApplication.Options.ArabicNumeral = .ArabicNumeral
                End If
                '------------------------------------------------------------------------------------------
                If .ShowBookmarks.HasValue Then
                    If Me.WordApplication.ActiveWindow.View.Type <> Word.WdViewType.wdPrintPreview Then
                        Me.WordApplication.ActiveWindow.View.ShowBookmarks = .ShowBookmarks
                    End If
                End If
                '------------------------------------------------------------------------------------------
                If .ShowTableGridlines.HasValue Then
                    Me.WordApplication.ActiveWindow.View.TableGridlines = .ShowTableGridlines
                End If
                '------------------------------------------------------------------------------------------
                If .DisplayPageBoundaries.HasValue Then
                    If Me.WordApplication.ActiveWindow.View.Type <> Word.WdViewType.wdPrintPreview Then
                        Me.WordApplication.ActiveWindow.View.DisplayPageBoundaries = .DisplayPageBoundaries
                    End If
                End If
                '------------------------------------------------------------------------------------------
                If .NormalViewZoomPercentage.HasValue Then
                    Me.WordApplication.ActiveWindow.View.Zoom.Percentage = .NormalViewZoomPercentage
                End If
                '----------------------------------------------------------------------------------
                If .NormalViewZoomPageFit.HasValue Then
                    Me.WordApplication.ActiveWindow.View.Zoom.PageFit = .NormalViewZoomPageFit
                End If
                '------------------------------------------------------------------------------------------
                If .NormalViewDisplayRulers.HasValue Then
                    Me.WordApplication.ActiveWindow.DisplayRulers = .NormalViewDisplayRulers
                End If

                Try
                    If Me.WordApplication.ActiveWindow.Panes.Count > 1 Then
                        Me.WordApplication.ActiveWindow.ActivePane.Close()
                    End If
                Catch ex As Exception
                End Try


                '----- Print Preview Section --------------------------------------------------------------
                If .ViewType.HasValue Then
                    Me.WordApplication.ActiveWindow.View.Type = .ViewType
                    If .ViewType = Word.WdViewType.wdPrintPreview Then
                        '----------------------------------------------------------------------------------
                        If .PrintPreviewZoomPercentage.HasValue Then
                            Me.WordApplication.ActiveWindow.View.Zoom.Percentage = .PrintPreviewZoomPercentage
                        End If
                        '----------------------------------------------------------------------------------
                        If .PrintPreviewZoomPageFit.HasValue Then
                            Me.WordApplication.ActiveWindow.View.Zoom.PageFit = .PrintPreviewZoomPageFit
                        End If
                        '----------------------------------------------------------------------------------
                        If .PrintPreviewPageFitness IsNot Nothing Then
                            Me.SetPrintPreviewPageFitness(.PrintPreviewPageFitness.Columns, _
                                                          .PrintPreviewPageFitness.Rows)
                        End If
                        '----------------------------------------------------------------------------------
                        If .PrintPreviewDisplayRulers.HasValue Then
                            Me.WordApplication.ActiveWindow.DisplayRulers = .PrintPreviewDisplayRulers
                        End If
                        '----------------------------------------------------------------------------------
                    End If
                End If
                '------------------------------------------------------------------------------------------
            End With
        End Sub


    End Class

End Namespace