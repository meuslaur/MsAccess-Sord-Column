Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15307
    DatasheetFontHeight =11
    ItemSuffix =16
    Left =645
    Top =450
    Right =16680
    Bottom =7875
    RecSrcDt = Begin
        0x9dd48382ced2e540
    End
    RecordSource ="T_Info"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =460
            Name ="EntêteFormulaire"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Width =1395
                    Height =460
                    FontSize =18
                    Name ="Auto_EnTete0"
                    Caption ="~T_Info"
                    FontName ="Calibri Light"
                    GridlineColor =10921638
                    LayoutCachedWidth =1395
                    LayoutCachedHeight =460
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =2324
            Name ="Détail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =113
                    Top =56
                    Width =2145
                    Height =360
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID_Info"
                    ControlSource ="ID_Info"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =113
                    LayoutCachedTop =56
                    LayoutCachedWidth =2258
                    LayoutCachedHeight =416
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1700
                    Top =510
                    Width =7950
                    Height =1715
                    TabIndex =2
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Info"
                    ControlSource ="InfoTexte"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1700
                    LayoutCachedTop =510
                    LayoutCachedWidth =9650
                    LayoutCachedHeight =2225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =340
                            Top =510
                            Width =1290
                            Height =1035
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Étiquette3"
                            Caption ="InfoTexte"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =340
                            LayoutCachedTop =510
                            LayoutCachedWidth =1630
                            LayoutCachedHeight =1545
                            RowStart =1
                            RowEnd =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3455
                    Top =56
                    Width =6186
                    Height =315
                    TabIndex =1
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="InfoTitre"
                    ControlSource ="InfoTitre"
                    GridlineColor =10921638

                    LayoutCachedLeft =3455
                    LayoutCachedTop =56
                    LayoutCachedWidth =9641
                    LayoutCachedHeight =371
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2324
                            Top =56
                            Width =1020
                            Height =315
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Étiquette10"
                            Caption ="InfoTitre"
                            GridlineColor =10921638
                            LayoutCachedLeft =2324
                            LayoutCachedTop =56
                            LayoutCachedWidth =3344
                            LayoutCachedHeight =371
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9807
                    Top =566
                    Width =4935
                    Height =1698
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Code"
                    ControlSource ="Code"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =9807
                    LayoutCachedTop =566
                    LayoutCachedWidth =14742
                    LayoutCachedHeight =2264
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="PiedFormulaire"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
