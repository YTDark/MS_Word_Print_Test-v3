

Imports Microsoft.Office.Interop

Namespace MSO

    Namespace Enums

        Public Module MSOEnums

            Public Enum MSFileFormat As Integer
                Word_2000_2003_2007_Suported = 100
                Document = Word.WdSaveFormat.wdFormatDocument
                RTF = Word.WdSaveFormat.wdFormatRTF
                Template = Word.WdSaveFormat.wdFormatTemplate
                HTML = Word.WdSaveFormat.wdFormatHTML
                Text = Word.WdSaveFormat.wdFormatText
                UnicodeText = Word.WdSaveFormat.wdFormatUnicodeText
                XML = Word.WdSaveFormat.wdFormatXML
                PDF = Word.WdSaveFormat.wdFormatPDF
            End Enum

            Public Enum MSWindowState As Integer
                Maximize = Word.WdWindowState.wdWindowStateMaximize
                Minimize = Word.WdWindowState.wdWindowStateMinimize
                Normal = Word.WdWindowState.wdWindowStateNormal
            End Enum

            Public Enum MSViewType As Integer
                MasterView = Word.WdViewType.wdMasterView
                NormalView = Word.WdViewType.wdNormalView
                OutlineView = Word.WdViewType.wdOutlineView
                PrintPreview = Word.WdViewType.wdPrintPreview
                PrintView = Word.WdViewType.wdPrintView
                ReadingView = Word.WdViewType.wdReadingView
                WebView = Word.WdViewType.wdWebView
            End Enum

            Public Enum MSArabicNumeral As Integer
                NumeralArabic = Word.WdArabicNumeral.wdNumeralArabic
                NumeralContext = Word.WdArabicNumeral.wdNumeralContext
                NumeralHindi = Word.WdArabicNumeral.wdNumeralHindi
                NumeralSystem = Word.WdArabicNumeral.wdNumeralSystem
            End Enum

            Public Enum MSPageFit As Integer
                PageFitBestFit = Word.WdPageFit.wdPageFitBestFit
                PageFitFullPage = Word.WdPageFit.wdPageFitFullPage
                PageFitNone = Word.WdPageFit.wdPageFitNone
                PageFitTextFit = Word.WdPageFit.wdPageFitTextFit
            End Enum

            Public Enum MSColor As Integer
                Aqua = Word.WdColor.wdColorAqua
                Automatic = Word.WdColor.wdColorAutomatic
                Black = Word.WdColor.wdColorBlack
                Blue = Word.WdColor.wdColorBlue
                BlueGray = Word.WdColor.wdColorBlueGray
                BrightGreen = Word.WdColor.wdColorBrightGreen
                Brown = Word.WdColor.wdColorBrown
                DarkBlue = Word.WdColor.wdColorDarkBlue
                DarkGreen = Word.WdColor.wdColorDarkGreen
                DarkRed = Word.WdColor.wdColorDarkRed
                DarkTeal = Word.WdColor.wdColorDarkTeal
                DarkYellow = Word.WdColor.wdColorDarkYellow
                Gold = Word.WdColor.wdColorGold
                Gray05 = Word.WdColor.wdColorGray05
                Gray10 = Word.WdColor.wdColorGray10
                Gray125 = Word.WdColor.wdColorGray125
                Gray15 = Word.WdColor.wdColorGray15
                Gray20 = Word.WdColor.wdColorGray20
                Gray25 = Word.WdColor.wdColorGray25
                Gray30 = Word.WdColor.wdColorGray30
                Gray35 = Word.WdColor.wdColorGray35
                Gray375 = Word.WdColor.wdColorGray375
                Gray40 = Word.WdColor.wdColorGray40
                Gray45 = Word.WdColor.wdColorGray45
                Gray50 = Word.WdColor.wdColorGray50
                Gray55 = Word.WdColor.wdColorGray55
                Gray60 = Word.WdColor.wdColorGray60
                Gray625 = Word.WdColor.wdColorGray625
                Gray65 = Word.WdColor.wdColorGray65
                Gray70 = Word.WdColor.wdColorGray70
                Gray75 = Word.WdColor.wdColorGray75
                Gray80 = Word.WdColor.wdColorGray80
                Gray85 = Word.WdColor.wdColorGray85
                Gray875 = Word.WdColor.wdColorGray875
                Gray90 = Word.WdColor.wdColorGray90
                Gray95 = Word.WdColor.wdColorGray95
                Green = Word.WdColor.wdColorGreen
                Indigo = Word.WdColor.wdColorIndigo
                Lavender = Word.WdColor.wdColorLavender
                LightBlue = Word.WdColor.wdColorLightBlue
                LightGreen = Word.WdColor.wdColorLightGreen
                LightOrange = Word.WdColor.wdColorLightOrange
                LightTurquoise = Word.WdColor.wdColorLightTurquoise
                LightYellow = Word.WdColor.wdColorLightYellow
                Lime = Word.WdColor.wdColorLime
                OliveGreen = Word.WdColor.wdColorOliveGreen
                Orange = Word.WdColor.wdColorOrange
                PaleBlue = Word.WdColor.wdColorPaleBlue
                Pink = Word.WdColor.wdColorPink
                Plum = Word.WdColor.wdColorPlum
                Red = Word.WdColor.wdColorRed
                Rose = Word.WdColor.wdColorRose
                SeaGreen = Word.WdColor.wdColorSeaGreen
                SkyBlue = Word.WdColor.wdColorSkyBlue
                Tan = Word.WdColor.wdColorTan
                Teal = Word.WdColor.wdColorTeal
                Turquoise = Word.WdColor.wdColorTurquoise
                Violet = Word.WdColor.wdColorViolet
                White = Word.WdColor.wdColorWhite
                Yellow = Word.WdColor.wdColorYellow
            End Enum

            Public Enum MSLineWidth As Integer
                LineWidth025pt = Word.WdLineWidth.wdLineWidth025pt
                LineWidth050pt = Word.WdLineWidth.wdLineWidth050pt
                LineWidth075pt = Word.WdLineWidth.wdLineWidth075pt
                LineWidth100pt = Word.WdLineWidth.wdLineWidth100pt
                LineWidth150pt = Word.WdLineWidth.wdLineWidth150pt
                LineWidth225pt = Word.WdLineWidth.wdLineWidth225pt
                LineWidth300pt = Word.WdLineWidth.wdLineWidth300pt
                LineWidth450pt = Word.WdLineWidth.wdLineWidth450pt
                LineWidth600pt = Word.WdLineWidth.wdLineWidth600pt
            End Enum

            Public Enum MSLineStyle As Integer
                DashDot = Word.WdLineStyle.wdLineStyleDashDot
                DashDotDot = Word.WdLineStyle.wdLineStyleDashDotDot
                DashDotStroked = Word.WdLineStyle.wdLineStyleDashDotStroked
                DashLargeGap = Word.WdLineStyle.wdLineStyleDashLargeGap
                DashSmallGap = Word.WdLineStyle.wdLineStyleDashSmallGap
                Dot = Word.WdLineStyle.wdLineStyleDot
                [Double] = Word.WdLineStyle.wdLineStyleDouble
                DoubleWavy = Word.WdLineStyle.wdLineStyleDoubleWavy
                Emboss3D = Word.WdLineStyle.wdLineStyleEmboss3D
                Engrave3D = Word.WdLineStyle.wdLineStyleEngrave3D
                Inset = Word.WdLineStyle.wdLineStyleInset
                None = Word.WdLineStyle.wdLineStyleNone
                Outset = Word.WdLineStyle.wdLineStyleOutset
                [Single] = Word.WdLineStyle.wdLineStyleSingle
                SingleWavy = Word.WdLineStyle.wdLineStyleSingleWavy
                ThickThinLargeGap = Word.WdLineStyle.wdLineStyleThickThinLargeGap
                ThickThinMedGap = Word.WdLineStyle.wdLineStyleThickThinMedGap
                ThickThinSmallGap = Word.WdLineStyle.wdLineStyleThickThinSmallGap
                ThinThickLargeGap = Word.WdLineStyle.wdLineStyleThinThickLargeGap
                ThinThickMedGap = Word.WdLineStyle.wdLineStyleThinThickMedGap
                ThinThickSmallGap = Word.WdLineStyle.wdLineStyleThinThickSmallGap
                ThinThickThinLargeGap = Word.WdLineStyle.wdLineStyleThinThickThinLargeGap
                ThinThickThinMedGap = Word.WdLineStyle.wdLineStyleThinThickThinMedGap
                ThinThickThinSmallGap = Word.WdLineStyle.wdLineStyleThinThickThinSmallGap
            End Enum

            Public Enum ImageSource
                BytesArray = 0
                FilePath = 1
            End Enum

        End Module

    End Namespace

End Namespace