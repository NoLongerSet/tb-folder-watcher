Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =9
    Right =50000
    Bottom =50000
    RecSrcDt = Begin
        0x0eb1fbebcc85e640
    End
    Caption ="Folder Watcher"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Aptos"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Aptos"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            TextFontFamily =0
            FontSize =11
            FontWeight =400
            FontName ="Aptos"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Aptos"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ToggleButton
            TextFontFamily =0
            FontName ="Aptos"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =1
        End
        Begin Section
            Height =8385
            BackColor =3621102
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =1380
                    Top =840
                    Width =6840
                    Height =2400
                    FontSize =48
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label0"
                    Caption ="twinBASIC \015\012Folder Watcher"
                    FontName ="Aptos Display"
                    LayoutCachedLeft =1380
                    LayoutCachedTop =840
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =3240
                    ThemeFontIndex =0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =4140
                    Width =6600
                    Height =540
                    FontSize =16
                    Name ="tbFolderToWatch"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =4140
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =4680
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1560
                            Top =3660
                            Width =3000
                            Height =435
                            FontSize =16
                            FontWeight =700
                            ForeColor =16777215
                            Name ="Label2"
                            Caption ="Folder to Watch:"
                            LayoutCachedLeft =1560
                            LayoutCachedTop =3660
                            LayoutCachedWidth =4560
                            LayoutCachedHeight =4095
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =1560
                    Top =4920
                    Width =6540
                    Height =1560
                    FontSize =20
                    FontWeight =400
                    TabIndex =1
                    Name ="toggleWatchFolder"
                    AfterUpdate ="[Event Procedure]"
                    Caption ="status: NOT Watching\015\012(press to start)"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =4920
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =6480
                    Shape =1
                    Bevel =0
                    Gradient =12
                    OldBorderStyle =1
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1620
                    Top =7020
                    Width =4140
                    Height =585
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    ForeColor =16777215
                    Name ="tbVersion"
                    ControlSource ="=\"Version \" & GetAppVersion()"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =7020
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =7605
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
    End
End
CodeBehindForm
' See "FolderWatcher.cls"
