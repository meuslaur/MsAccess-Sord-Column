Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    FastLaserPrinting = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9651
    DatasheetFontHeight =11
    ItemSuffix =50
    Left =2400
    Top =1185
    Right =12045
    Bottom =9735
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x4a3e621c3cd1e540
    End
    Caption ="Message d'information"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationControl
            BorderWidth =1
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationButton
            Width =283
            Height =283
            ForeColor =-2
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            HoverColor =-2
            HoverThemeColorIndex =2
            HoverTint =20.0
            PressedColor =-2
            PressedThemeColorIndex =2
            PressedTint =60.0
            HoverForeColor =-2
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =-2
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            BackThemeColorIndex =1
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            FontName ="Calibri"
            FontWeight =400
            FontSize =11
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Section
            Height =8560
            BackColor =15921906
            Name ="Détail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =9651
                    Height =503
                    BorderColor =10921638
                    Name ="boxTitre"
                    GridlineColor =10921638
                    LayoutCachedWidth =9651
                    LayoutCachedHeight =503
                    BackThemeColorIndex =6
                    BackTint =40.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =7937
                    Width =9651
                    Height =623
                    BorderColor =10921638
                    Name ="boxBouton"
                    GridlineColor =10921638
                    LayoutCachedTop =7937
                    LayoutCachedWidth =9651
                    LayoutCachedHeight =8560
                    BackThemeColorIndex =6
                    BackTint =40.0
                End
                Begin CommandButton
                    Default = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =2494
                    Top =8050
                    Width =1710
                    Height =405
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmbValider"
                    Caption ="Valider"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2494
                    LayoutCachedTop =8050
                    LayoutCachedWidth =4204
                    LayoutCachedHeight =8455
                    UseTheme =0
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =118
                    Top =680
                    Width =2145
                    Height =300
                    BackColor =5855577
                    BorderColor =8355711
                    Name ="lbl_Bdd"
                    Caption ="Base de données"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =118
                    LayoutCachedTop =680
                    LayoutCachedWidth =2263
                    LayoutCachedHeight =980
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =5329
                    Width =2145
                    Height =300
                    BorderColor =8355711
                    Name ="lbl_TxtBox"
                    Caption ="TextBox"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =5329
                    LayoutCachedWidth =2265
                    LayoutCachedHeight =5629
                    ThemeFontIndex =-1
                    BackThemeColorIndex =7
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =113
                    Top =5782
                    Width =2145
                    Height =300
                    BorderColor =8355711
                    Name ="lbl_Cmb"
                    Caption ="CommandButton"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =5782
                    LayoutCachedWidth =2258
                    LayoutCachedHeight =6082
                    ThemeFontIndex =-1
                    BackThemeColorIndex =9
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =113
                    Top =4195
                    Width =2160
                    Height =300
                    BorderColor =8355711
                    Name ="lbl_Code"
                    Caption ="Code"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =4195
                    LayoutCachedWidth =2273
                    LayoutCachedHeight =4495
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =113
                    Top =1700
                    Width =2145
                    Height =300
                    BorderColor =8355711
                    Name ="lbl_Form"
                    Caption ="Formulaire"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =1700
                    LayoutCachedWidth =2258
                    LayoutCachedHeight =2000
                    ThemeFontIndex =-1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =118
                    Top =2267
                    Width =2145
                    Height =300
                    BackColor =4138256
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lbl_Source"
                    Caption ="Source"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =118
                    LayoutCachedTop =2267
                    LayoutCachedWidth =2263
                    LayoutCachedHeight =2567
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BackShade =50.0
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =113
                    Top =2834
                    Width =2145
                    Height =300
                    BackColor =9592887
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lbl_Fields"
                    Caption ="Champs"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =2834
                    LayoutCachedWidth =2258
                    LayoutCachedHeight =3134
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5385
                    Top =8050
                    Width =1710
                    Height =405
                    ForeColor =4210752
                    Name ="cmbAnnuler"
                    Caption ="Annuler"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5385
                    LayoutCachedTop =8050
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =8455
                    UseTheme =0
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =113
                    Top =56
                    Width =9411
                    Height =360
                    FontSize =13
                    ForeColor =-2147483641
                    Name ="lbl_Titre"
                    Caption ="Résumé des opérations à effectuer"
                    GridlineColor =10921638
                    LayoutCachedLeft =113
                    LayoutCachedTop =56
                    LayoutCachedWidth =9524
                    LayoutCachedHeight =416
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =2381
                    Top =2834
                    Width =7140
                    Height =855
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lbl_Champs"
                    Caption ="lbl_Champs"
                    GridlineColor =10921638
                    LayoutCachedLeft =2381
                    LayoutCachedTop =2834
                    LayoutCachedWidth =9521
                    LayoutCachedHeight =3689
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2379
                    Top =6236
                    Width =6690
                    Height =315
                    TabIndex =2
                    ForeColor =6710886
                    Name ="txtPicFold"
                    GridlineColor =10921638

                    LayoutCachedLeft =2379
                    LayoutCachedTop =6236
                    LayoutCachedWidth =9069
                    LayoutCachedHeight =6551
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =565
                            Top =6236
                            Width =1700
                            Height =341
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtPicFold"
                            Caption ="Dossier img"
                            GridlineColor =10921638
                            LayoutCachedLeft =565
                            LayoutCachedTop =6236
                            LayoutCachedWidth =2265
                            LayoutCachedHeight =6577
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2379
                    Top =6689
                    Width =3855
                    Height =315
                    TabIndex =3
                    ForeColor =6710886
                    Name ="txtPicAsc"
                    GridlineColor =10921638

                    LayoutCachedLeft =2379
                    LayoutCachedTop =6689
                    LayoutCachedWidth =6234
                    LayoutCachedHeight =7004
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2379
                    Top =7086
                    Width =3855
                    Height =315
                    TabIndex =4
                    ForeColor =6710886
                    Name ="txtPicDesc"
                    GridlineColor =10921638

                    LayoutCachedLeft =2379
                    LayoutCachedTop =7086
                    LayoutCachedWidth =6234
                    LayoutCachedHeight =7401
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =563
                            Top =7086
                            Width =1695
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtPicDesc"
                            Caption ="Img DESC"
                            GridlineColor =10921638
                            LayoutCachedLeft =563
                            LayoutCachedTop =7086
                            LayoutCachedWidth =2258
                            LayoutCachedHeight =7401
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3683
                    Top =5329
                    Width =1125
                    Height =315
                    TabIndex =5
                    ForeColor =6710886
                    Name ="txtTbSuffixe"
                    GridlineColor =10921638

                    LayoutCachedLeft =3683
                    LayoutCachedTop =5329
                    LayoutCachedWidth =4808
                    LayoutCachedHeight =5644
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3683
                    Top =5782
                    Width =1125
                    Height =315
                    TabIndex =6
                    ForeColor =6710886
                    Name ="txtCmbSuffixe"
                    GridlineColor =10921638

                    LayoutCachedLeft =3683
                    LayoutCachedTop =5782
                    LayoutCachedWidth =4808
                    LayoutCachedHeight =6097
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2379
                    Top =5782
                    Width =1140
                    Height =315
                    TabIndex =7
                    ForeColor =6710886
                    Name ="txtCmbPrefixe"
                    GridlineColor =10921638

                    LayoutCachedLeft =2379
                    LayoutCachedTop =5782
                    LayoutCachedWidth =3519
                    LayoutCachedHeight =6097
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2379
                    Top =5329
                    Width =1140
                    Height =315
                    TabIndex =8
                    ForeColor =6710886
                    Name ="txtTbPrefixe"
                    GridlineColor =10921638

                    LayoutCachedLeft =2379
                    LayoutCachedTop =5329
                    LayoutCachedWidth =3519
                    LayoutCachedHeight =5644
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2367
                    Top =4195
                    Width =1650
                    Height =315
                    TabIndex =9
                    ForeColor =6710886
                    Name ="txtCodeVar"
                    GridlineColor =10921638

                    LayoutCachedLeft =2367
                    LayoutCachedTop =4195
                    LayoutCachedWidth =4017
                    LayoutCachedHeight =4510
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            Left =2381
                            Top =3798
                            Width =1020
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtCodeVar"
                            Caption ="Var Classe"
                            GridlineColor =10921638
                            LayoutCachedLeft =2381
                            LayoutCachedTop =3798
                            LayoutCachedWidth =3401
                            LayoutCachedHeight =4113
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4251
                    Top =4195
                    Width =1710
                    Height =315
                    TabIndex =10
                    ForeColor =6710886
                    Name ="txtCodeFunc"
                    GridlineColor =10921638

                    LayoutCachedLeft =4251
                    LayoutCachedTop =4195
                    LayoutCachedWidth =5961
                    LayoutCachedHeight =4510
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            Left =4251
                            Top =3798
                            Width =1425
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtCodeFunc"
                            Caption ="Fonction"
                            GridlineColor =10921638
                            LayoutCachedLeft =4251
                            LayoutCachedTop =3798
                            LayoutCachedWidth =5676
                            LayoutCachedHeight =4113
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2381
                    Top =2267
                    Width =4980
                    Height =315
                    TabIndex =11
                    ForeColor =6710886
                    Name ="txtFrmSource"
                    GridlineColor =10921638

                    LayoutCachedLeft =2381
                    LayoutCachedTop =2267
                    LayoutCachedWidth =7361
                    LayoutCachedHeight =2582
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2381
                    Top =1700
                    Width =4980
                    Height =315
                    TabIndex =12
                    ForeColor =6710886
                    Name ="txtFrmNom"
                    GridlineColor =10921638

                    LayoutCachedLeft =2381
                    LayoutCachedTop =1700
                    LayoutCachedWidth =7361
                    LayoutCachedHeight =2015
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2381
                    Top =1077
                    Width =7140
                    Height =315
                    FontWeight =700
                    TabIndex =13
                    ForeColor =6710886
                    Name ="txtBddSauve"
                    GridlineColor =10921638

                    LayoutCachedLeft =2381
                    LayoutCachedTop =1077
                    LayoutCachedWidth =9521
                    LayoutCachedHeight =1392
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =567
                            Top =1077
                            Width =1695
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtBddSauve"
                            Caption ="Sauvegarde"
                            GridlineColor =10921638
                            LayoutCachedLeft =567
                            LayoutCachedTop =1077
                            LayoutCachedWidth =2262
                            LayoutCachedHeight =1392
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2381
                    Top =680
                    Width =7140
                    Height =315
                    TabIndex =14
                    ForeColor =6710886
                    Name ="txtBdd"
                    GridlineColor =10921638

                    LayoutCachedLeft =2381
                    LayoutCachedTop =680
                    LayoutCachedWidth =9521
                    LayoutCachedHeight =995
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =60.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =561
                    Top =6689
                    Width =1695
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lbl_txtPicAsc"
                    Caption ="Img ASC"
                    GridlineColor =10921638
                    LayoutCachedLeft =561
                    LayoutCachedTop =6689
                    LayoutCachedWidth =2256
                    LayoutCachedHeight =7004
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2381
                    Top =4988
                    Width =1125
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lbl_Prefixe"
                    Caption ="Préfixe"
                    GridlineColor =10921638
                    LayoutCachedLeft =2381
                    LayoutCachedTop =4988
                    LayoutCachedWidth =3506
                    LayoutCachedHeight =5303
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =3685
                    Top =4988
                    Width =1125
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lbl_Suffixe"
                    Caption ="Suffixe"
                    GridlineColor =10921638
                    LayoutCachedLeft =3685
                    LayoutCachedTop =4988
                    LayoutCachedWidth =4810
                    LayoutCachedHeight =5303
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =737
                    Top =7766
                    Width =680
                    Height =341
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Étiquette46"
                    Caption ="j"
                    GridlineColor =10921638
                    LayoutCachedLeft =737
                    LayoutCachedTop =7766
                    LayoutCachedWidth =1417
                    LayoutCachedHeight =8107
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    
    If IsNull(Me.OpenArgs) Then Exit Sub

    With Forms.Item("F_CreateForm")
        txtBdd = .Controls("txtBdd").Value
        txtBddSauve = .Controls("txtBddSauve").Value
        txtFrmNom = .Controls("txtFormName").Value
        txtFrmSource = .Controls("txtFormSource").Value
        lbl_Champs.Caption = .Controls("txtCacheFields").Value
        txtCodeVar = .Controls("txtClasseName").Value
        txtCodeFunc = .Controls("txtFunctionName").Value
        txtTbPrefixe = Nz(.Controls("txtTbPrefix").Value, "")
        txtTbSuffixe = Nz(.Controls("txtTbSuffix").Value, "")
        txtCmbPrefixe = Nz(.Controls("txtCmbPrefix").Value, "")
        txtCmbSuffixe = Nz(.Controls("txtCmbSuffix").Value, "")
        txtPicFold = Nz(.Controls("txtPicFolder").Value, "")
        txtPicAsc = Nz(.Controls("txtPicAsc").Value, "")
        txtPicDesc = Nz(.Controls("txtPicDesc").Value, "")
    End With

End Sub

Private Sub cmbAnnuler_Click()
    Forms.Item("F_CreateForm").Controls("txtCacheFields") = "ANNULER"
    DoCmd.Close
End Sub

Private Sub cmbValider_Click()
    DoCmd.Close
End Sub
