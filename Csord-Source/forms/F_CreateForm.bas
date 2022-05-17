Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    FastLaserPrinting = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14173
    DatasheetFontHeight =11
    ItemSuffix =104
    Left =5610
    Top =225
    Right =19785
    Bottom =11655
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x80c066f2cdd0e540
    End
    Caption ="Paramètrage pour la création du formulaire"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnTimer ="=Raz_bErr()"
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
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
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ToggleButton
            Width =283
            Height =283
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin WebBrowser
            OldBorderStyle =1
            Width =4536
            Height =2835
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="F_Entete"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            Height =11451
            Name ="F_Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =5840
                    Top =3741
                    Width =7760
                    Height =1191
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="boxForm"
                    GridlineColor =10921638
                    LayoutCachedLeft =5840
                    LayoutCachedTop =3741
                    LayoutCachedWidth =13600
                    LayoutCachedHeight =4932
                    BackShade =95.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =6464
                    Top =8335
                    Width =6802
                    Height =2829
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="boxCmb"
                    GridlineColor =10921638
                    LayoutCachedLeft =6464
                    LayoutCachedTop =8335
                    LayoutCachedWidth =13266
                    LayoutCachedHeight =11164
                    BackShade =95.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =6464
                    Top =7030
                    Width =6800
                    Height =1191
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="boxTxtBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =6464
                    LayoutCachedTop =7030
                    LayoutCachedWidth =13264
                    LayoutCachedHeight =8221
                    BackShade =95.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =6461
                    Top =5726
                    Width =6800
                    Height =1191
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="boxCode"
                    GridlineColor =10921638
                    LayoutCachedLeft =6461
                    LayoutCachedTop =5726
                    LayoutCachedWidth =13261
                    LayoutCachedHeight =6917
                    BackShade =95.0
                End
                Begin ListBox
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =745
                    Top =1458
                    Width =4536
                    Height =4095
                    TabIndex =8
                    BoundColumn =1
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lstObjets"
                    RowSourceType ="Value List"
                    RowSource ="6;T_Depenses;Table liée;6;T_Vehicules;Table liée;"
                    ColumnWidths ="0;2835;1134"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =745
                    LayoutCachedTop =1458
                    LayoutCachedWidth =5281
                    LayoutCachedHeight =5553
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =745
                            Top =1148
                            Width =4530
                            Height =315
                            BackColor =15921906
                            ForeColor =6710886
                            Name ="lbl_lstObjets"
                            Caption ="Objets de la base :"
                            GridlineColor =10921638
                            LayoutCachedLeft =745
                            LayoutCachedTop =1148
                            LayoutCachedWidth =5275
                            LayoutCachedHeight =1463
                            BackShade =95.0
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =65.0
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =7544
                    Top =4025
                    Width =4536
                    Height =315
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtFormName"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    ControlTipText ="Nom du formulaire"
                    GridlineColor =10921638

                    LayoutCachedLeft =7544
                    LayoutCachedTop =4025
                    LayoutCachedWidth =12080
                    LayoutCachedHeight =4340
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =6522
                            Top =4025
                            Width =930
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtFormName"
                            Caption ="Nom :"
                            GridlineColor =10921638
                            LayoutCachedLeft =6522
                            LayoutCachedTop =4025
                            LayoutCachedWidth =7452
                            LayoutCachedHeight =4340
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8853
                    Top =7258
                    Width =861
                    Height =315
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtTbPrefix"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"txtTb\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =8853
                    LayoutCachedTop =7258
                    LayoutCachedWidth =9714
                    LayoutCachedHeight =7573
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =7713
                            Top =7258
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtTbPrefix"
                            Caption ="Prefixe :"
                            GridlineColor =10921638
                            LayoutCachedLeft =7713
                            LayoutCachedTop =7258
                            LayoutCachedWidth =8793
                            LayoutCachedHeight =7573
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8857
                    Top =7654
                    Width =861
                    Height =315
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtTbSuffix"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"txtTb\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =8857
                    LayoutCachedTop =7654
                    LayoutCachedWidth =9718
                    LayoutCachedHeight =7969
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =7710
                            Top =7654
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtTbSuffix"
                            Caption ="Suffix :"
                            GridlineColor =10921638
                            LayoutCachedLeft =7710
                            LayoutCachedTop =7654
                            LayoutCachedWidth =8790
                            LayoutCachedHeight =7969
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8861
                    Top =8449
                    Width =846
                    Height =345
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtCmbPrefix"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"txtCmb\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =8861
                    LayoutCachedTop =8449
                    LayoutCachedWidth =9707
                    LayoutCachedHeight =8794
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =7710
                            Top =8448
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtCmbPrefix"
                            Caption ="Prefixe  :"
                            GridlineColor =10921638
                            LayoutCachedLeft =7710
                            LayoutCachedTop =8448
                            LayoutCachedWidth =8790
                            LayoutCachedHeight =8763
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8861
                    Top =8845
                    Width =846
                    Height =345
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtCmbSuffix"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"txtCmb\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =8861
                    LayoutCachedTop =8845
                    LayoutCachedWidth =9707
                    LayoutCachedHeight =9190
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =7710
                            Top =8845
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtCmbSuffix"
                            Caption ="Suffix  :"
                            GridlineColor =10921638
                            LayoutCachedLeft =7710
                            LayoutCachedTop =8845
                            LayoutCachedWidth =8790
                            LayoutCachedHeight =9160
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8857
                    Top =10205
                    Width =2901
                    Height =345
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtPicAsc"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"txtPic\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =8857
                    LayoutCachedTop =10205
                    LayoutCachedWidth =11758
                    LayoutCachedHeight =10550
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =6636
                            Top =10205
                            Width =2160
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtPicAsc"
                            Caption ="Image tri ASC :"
                            GridlineColor =10921638
                            LayoutCachedLeft =6636
                            LayoutCachedTop =10205
                            LayoutCachedWidth =8796
                            LayoutCachedHeight =10520
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8857
                    Top =10659
                    Width =2901
                    Height =345
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtPicDesc"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"txtPic\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =8857
                    LayoutCachedTop =10659
                    LayoutCachedWidth =11758
                    LayoutCachedHeight =11004
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =6639
                            Top =10659
                            Width =2160
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtPicDesc"
                            Caption ="Image tri DESC :"
                            GridlineColor =10921638
                            LayoutCachedLeft =6639
                            LayoutCachedTop =10659
                            LayoutCachedWidth =8799
                            LayoutCachedHeight =10974
                        End
                    End
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =5953
                    Top =7030
                    Width =345
                    Height =1200
                    FontSize =12
                    BorderColor =8355711
                    Name ="lbl_boxTxtBox"
                    Caption ="TextBox"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =5953
                    LayoutCachedTop =7030
                    LayoutCachedWidth =6298
                    LayoutCachedHeight =8230
                    ThemeFontIndex =-1
                    BackThemeColorIndex =7
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =5953
                    Top =8334
                    Width =330
                    Height =2820
                    FontSize =12
                    BorderColor =8355711
                    Name ="lbl_boxCmb"
                    Caption ="CommandButton"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =5953
                    LayoutCachedTop =8334
                    LayoutCachedWidth =6283
                    LayoutCachedHeight =11154
                    ThemeFontIndex =-1
                    BackThemeColorIndex =9
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8853
                    Top =6010
                    Height =315
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtClasseName"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    ControlTipText ="Nom variable de classe"
                    GridlineColor =10921638

                    LayoutCachedLeft =8853
                    LayoutCachedTop =6010
                    LayoutCachedWidth =10554
                    LayoutCachedHeight =6325
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =7149
                            Top =6010
                            Width =1635
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtClasseName"
                            Caption ="Variable classe :"
                            GridlineColor =10921638
                            LayoutCachedLeft =7149
                            LayoutCachedTop =6010
                            LayoutCachedWidth =8784
                            LayoutCachedHeight =6325
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8853
                    Top =6463
                    Height =315
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtFunctionName"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    ControlTipText ="Nom de la fonction"
                    GridlineColor =10921638

                    LayoutCachedLeft =8853
                    LayoutCachedTop =6463
                    LayoutCachedWidth =10554
                    LayoutCachedHeight =6778
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =7160
                            Top =6464
                            Width =1635
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtFunctionName"
                            Caption ="Nom Function :"
                            GridlineColor =10921638
                            LayoutCachedLeft =7160
                            LayoutCachedTop =6464
                            LayoutCachedWidth =8795
                            LayoutCachedHeight =6779
                        End
                    End
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =5950
                    Top =5726
                    Width =375
                    Height =1185
                    FontSize =12
                    BorderColor =8355711
                    Name ="lbl_boxCode"
                    Caption ="Code"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =5950
                    LayoutCachedTop =5726
                    LayoutCachedWidth =6325
                    LayoutCachedHeight =6911
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =8857
                    Top =9752
                    Width =3501
                    Height =345
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtPicFolder"
                    AfterUpdate ="=RestaureLabelTxt()"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =8857
                    LayoutCachedTop =9752
                    LayoutCachedWidth =12358
                    LayoutCachedHeight =10097
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =6643
                            Top =9752
                            Width =2145
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtPicFolder"
                            Caption ="Sous dossier images :"
                            GridlineColor =10921638
                            LayoutCachedLeft =6643
                            LayoutCachedTop =9752
                            LayoutCachedWidth =8788
                            LayoutCachedHeight =10067
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =5669
                    Top =5385
                    Width =7753
                    Height =5949
                    BorderColor =10921638
                    Name ="boxOptions"
                    GridlineColor =10921638
                    LayoutCachedLeft =5669
                    LayoutCachedTop =5385
                    LayoutCachedWidth =13422
                    LayoutCachedHeight =11334
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =5782
                    Top =5215
                    Width =2340
                    Height =315
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lbl_boxOptions"
                    Caption ="Paramètres de création :"
                    GridlineColor =10921638
                    LayoutCachedLeft =5782
                    LayoutCachedTop =5215
                    LayoutCachedWidth =8122
                    LayoutCachedHeight =5530
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =12419
                    Top =9808
                    Width =345
                    Height =285
                    TabIndex =22
                    ForeColor =4210752
                    Name ="cmbSelectPicFolder"
                    Caption =",,,"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sélection du dossier..."
                    GridlineColor =10921638

                    LayoutCachedLeft =12419
                    LayoutCachedTop =9808
                    LayoutCachedWidth =12764
                    LayoutCachedHeight =10093
                    UseTheme =0
                    BackColor =14461583
                    BorderWidth =1
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
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =11805
                    Top =10715
                    Width =345
                    Height =285
                    TabIndex =24
                    ForeColor =4210752
                    Name ="cmbSelectPicDesc"
                    Caption =",,,"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sélection de l'image..."
                    GridlineColor =10921638

                    LayoutCachedLeft =11805
                    LayoutCachedTop =10715
                    LayoutCachedWidth =12150
                    LayoutCachedHeight =11000
                    UseTheme =0
                    BackColor =14461583
                    BorderWidth =1
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
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =11805
                    Top =10261
                    Width =345
                    Height =285
                    TabIndex =23
                    ForeColor =4210752
                    Name ="cmbSelectPicAsc"
                    Caption =",,,"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sélection de l'image..."
                    GridlineColor =10921638

                    LayoutCachedLeft =11805
                    LayoutCachedTop =10261
                    LayoutCachedWidth =12150
                    LayoutCachedHeight =10546
                    UseTheme =0
                    BackColor =14461583
                    BorderWidth =1
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
                Begin Image
                    Left =12753
                    Top =8417
                    Width =450
                    Height =300
                    BorderColor =10921638
                    Name ="img_InfoCommandButton"
                    Picture ="ic_CommandButton.png"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x89504e470d0a1a0a0000000d494844520000001e00000014080200000015c91a ,
                        0x93000000017352474200aece1ce90000000467414d410000b18f0bfc61050000 ,
                        0x00097048597300000ec300000ec301c76fa8640000005849444154484bedd4b1 ,
                        0x11c0300843d18c04ded99e2803e5fbc49d933ad0a153215cbcd2d7702beaa6ef ,
                        0xd4ac393fb48eff7deb253461374de51276d3542e61374de51276d3542e61079d ,
                        0x9ea0f903e3212f4157e887ceafdb03ecac938dab3464180000000049454e44ae ,
                        0x426082
                    End

                    LayoutCachedLeft =12753
                    LayoutCachedTop =8417
                    LayoutCachedWidth =13203
                    LayoutCachedHeight =8717
                    TabIndex =30
                End
                Begin Image
                    Left =12757
                    Top =7087
                    Width =450
                    Height =330
                    BorderColor =10921638
                    Name ="img_InfoTextBox"
                    Picture ="ic_TextBox.png"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x89504e470d0a1a0a0000000d494844520000001e0000001608020000005801bb ,
                        0x98000000017352474200aece1ce90000000467414d410000b18f0bfc61050000 ,
                        0x00097048597300000ec300000ec301c76fa864000001bf49444154484b633432 ,
                        0x3460a00d0019bd68f112288f7a202e36066a34900515a3068018c804e5510c30 ,
                        0xbd4eacd164041ad55c8d0986a6d12c501a0694949482828279f9f880eccf9f3e ,
                        0xcd9d3be7fdfbf7102920484d4b939191053230a53001baab81e63e78f060e182 ,
                        0xf940242a26969c9c02956060c8cdcd0392102920a3a2b20a2c8c13a01bddd3d3 ,
                        0xbd66cdea7b60b065cb66155555a80403c38f9f3f66cf9a0591022ae360670f09 ,
                        0x0985ca6103d8c35a5050d0d6ce4e5c4c9c8b8b0b2ac4c0b077cf1e280b0cae5c ,
                        0xbd22282408e56003e846031d3271e224a0675d5c5cd174021d0b65c10024dc71 ,
                        0x01946804bad4d6d676e2c4091053406eb7b5834861054f9e3c86b2b00114576b ,
                        0x686800bd09779d8585258401019e9e5e501618e868eb3c79fc04cac106508cfe ,
                        0xf1fd07dc8f40271b191b43d810606169091484b081a9f0c7cf9fdbb76f8370b1 ,
                        0x02940001268982c22260587ff8f891838363f9f265eaeaea50390686f5ebd7b5 ,
                        0x7774be7cf952809f1fa8a0a3bd0d2a8103a0180dcc02f575b506868640f685f3 ,
                        0xe78124bcb0853032d2d380f1f1f4c913cc28c50458121fd05088b958c1e14387 ,
                        0x88311708b0a76b3200666542e30a0ccaa32e606000009c12a9f9d6c057460000 ,
                        0x000049454e44ae426082
                    End

                    LayoutCachedLeft =12757
                    LayoutCachedTop =7087
                    LayoutCachedWidth =13207
                    LayoutCachedHeight =7417
                    TabIndex =29
                End
                Begin Image
                    Left =13097
                    Top =3627
                    Width =450
                    Height =360
                    BorderColor =10921638
                    Name ="img_InfoForm"
                    Picture ="ic_Form.png"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x89504e470d0a1a0a0000000d494844520000001e000000180802000000620bda ,
                        0xe8000000017352474200aece1ce90000000467414d410000b18f0bfc61050000 ,
                        0x00097048597300000ec300000ec301c76fa864000000dc49444154484b63d4d4 ,
                        0xd464a00d0019bd7af56a288f7a203434146a34900515a3068018c804e5d100d0 ,
                        0xd0e8e111d67c21adcc4232500e6500dd6816593dbeb00e764d47289f0280251a ,
                        0x1939f9b93d8a78bc4aa17c72010b94c6006cea767c5c029fd65443b8f9f9f910 ,
                        0x061e3071e244280b0c701afdebeea9ef47e6433918da88015802e4ffef9fdf8f ,
                        0x2ffbb2a9f9efbb275021b200bad17f5fdcfab4ace0fb89e5503e0500dde88fcb ,
                        0x8b29742c1ce00c6b34408568c43402ae017f4c626a44379a8c94800b101b2040 ,
                        0x404c982083d1aa000dd0b82a80f2a80b181800b861522ae6e5d8f50000000049 ,
                        0x454e44ae426082
                    End

                    LayoutCachedLeft =13097
                    LayoutCachedTop =3627
                    LayoutCachedWidth =13547
                    LayoutCachedHeight =3987
                    TabIndex =27
                End
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =1
                    Left =12810
                    Top =5787
                    Width =397
                    Height =397
                    BorderColor =10921638
                    Name ="img_InfoCode"
                    Picture ="ic_Code.png"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x89504e470d0a1a0a0000000d49484452000000200000001f0802000000093458 ,
                        0xdb0000000467414d410000b18f0bfc6105000000097048597300000ec300000e ,
                        0xc301c76fa864000001c049444154484bdd943b4b03411485fd3bdb884f0c592b ,
                        0x61d5c246411103a21054520c8414da042185a2b0458a5549a16011b0506c140b ,
                        0x058b14365b080928f8633cbb773233997d64573788f98a61e6dcc93db3772677 ,
                        0xc41830a90d5667d7ae0ad7add25b73f30673ae4693c2a0bc50b9ddba77d9fbf9 ,
                        0x7a03738c9843c19cef082391c1f1f2098efcbcfb8a0997ba40818e683044e806 ,
                        0xa38b1b33d5bb7cddc5883989edf207523c149f420da0238a3d58b67b81220d26 ,
                        0xb66bb9a397bcd399ae5c223546d840814e3fae2e1d522e1407d5c748aed011a5 ,
                        0x3d2ad2608a9d79b94e5b937bb61f9240813eef7c620f29488deb453aed92e30c ,
                        0xccc6171d36dc009f5577b1874b3e21e9e20d30aa251a2f1c7825723ab9da23e6 ,
                        0x628f20cac02bbc0214694078d5df6f42512f1964f0052a7d152d5d716e270303 ,
                        0xf5998a74e299e2519122486da0e6828178a6c13f07f1c312d13345eabebde857 ,
                        0x779084ff6370d10b573583b195922f72b47488c61844210d82bd48a413ad42f4 ,
                        0xa2e44803426d1530e0adc2efa67c474a7403d3343152abc09145abb02c8b3166 ,
                        0xdb36af71b7ca7c1180a24037a075860cb781564d5a06098d9208fef40b326128 ,
                        0x0d320769a5c140308c6f8b9e4a736714ddb00000000049454e44ae426082
                    End

                    LayoutCachedLeft =12810
                    LayoutCachedTop =5787
                    LayoutCachedWidth =13207
                    LayoutCachedHeight =6184
                    TabIndex =28
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5385
                    Top =3743
                    Width =375
                    Height =1185
                    FontSize =12
                    BorderColor =8355711
                    Name ="lbl_boxForm"
                    Caption ="Form"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =5385
                    LayoutCachedTop =3743
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4928
                    ThemeFontIndex =-1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5726
                    Top =1416
                    Width =8166
                    Height =2070
                    FontSize =10
                    TabIndex =6
                    BorderColor =2366701
                    ForeColor =4210752
                    Name ="txtInfoTxt"
                    GridlineColor =10921638

                    LayoutCachedLeft =5726
                    LayoutCachedTop =1416
                    LayoutCachedWidth =13892
                    LayoutCachedHeight =3486
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5499
                    Top =1133
                    Width =8391
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtInfoTitre"
                    GridlineColor =10921638

                    LayoutCachedLeft =5499
                    LayoutCachedTop =1133
                    LayoutCachedWidth =13890
                    LayoutCachedHeight =1418
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =295
                    Top =1133
                    Width =375
                    Height =4425
                    FontSize =12
                    BackColor =4138256
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lbl_lstObjetsInfo"
                    Caption ="Source"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =295
                    LayoutCachedTop =1133
                    LayoutCachedWidth =670
                    LayoutCachedHeight =5558
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BackShade =50.0
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =283
                    Top =5749
                    Width =375
                    Height =5475
                    FontSize =12
                    BackColor =9592887
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lbl_lstFieldsInfo"
                    Caption ="Champs de : T_Vehicules"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =283
                    LayoutCachedTop =5749
                    LayoutCachedWidth =658
                    LayoutCachedHeight =11224
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                End
                Begin ListBox
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    MultiSelect =1
                    IMESentenceMode =3
                    Left =737
                    Top =6066
                    Width =4551
                    Height =5175
                    TabIndex =9
                    BackColor =14610923
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lstFields"
                    RowSourceType ="Value List"
                    RowSource ="ID_Vehicule;Vehi_Marque;Vehi_Modele;ID_Etat;Vehi_KmActuel;ID_Vendeur;Vehi_DateAc"
                        "hat;Vehi_DatePremImmat;Vehi_DateImmat;Vehi_Prix;Vehi_KmAchat;Vehi_FichImage;Vehi"
                        "_Note;"
                    AfterUpdate ="=RestaureLabelLst()"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =737
                    LayoutCachedTop =6066
                    LayoutCachedWidth =5288
                    LayoutCachedHeight =11241
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =733
                            Top =5760
                            Width =4551
                            Height =315
                            BackColor =15921906
                            ForeColor =6710886
                            Name ="lbl_lstFields"
                            Caption ="Champs de la source : (sel. max 10 champs)"
                            GridlineColor =10921638
                            LayoutCachedLeft =733
                            LayoutCachedTop =5760
                            LayoutCachedWidth =5284
                            LayoutCachedHeight =6075
                            BackShade =95.0
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =7541
                    Top =4478
                    Width =4536
                    Height =315
                    TabIndex =11
                    BackColor =15921906
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtFormSource"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    ControlTipText ="Source du formulaire"
                    GridlineColor =10921638

                    LayoutCachedLeft =7541
                    LayoutCachedTop =4478
                    LayoutCachedWidth =12077
                    LayoutCachedHeight =4793
                    BackShade =95.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =6525
                            Top =4478
                            Width =930
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtFormSource"
                            Caption ="Source :"
                            GridlineColor =10921638
                            LayoutCachedLeft =6525
                            LayoutCachedTop =4478
                            LayoutCachedWidth =7455
                            LayoutCachedHeight =4793
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =226
                    Top =56
                    Width =13730
                    Height =906
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="boxBdD"
                    GridlineColor =10921638
                    LayoutCachedLeft =226
                    LayoutCachedTop =56
                    LayoutCachedWidth =13956
                    LayoutCachedHeight =962
                    BackShade =95.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2967
                    Top =169
                    Width =7461
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtBdd"
                    OnGotFocus ="=AfficheInfo(\"\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2967
                    LayoutCachedTop =169
                    LayoutCachedWidth =10428
                    LayoutCachedHeight =484
                    BackThemeColorIndex =-1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =223
                    Left =10485
                    Top =195
                    Width =345
                    Height =285
                    TabIndex =2
                    ForeColor =4210752
                    Name ="cmbSelectBdd"
                    Caption =",,,"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sélection de la base..."
                    GridlineColor =10921638

                    LayoutCachedLeft =10485
                    LayoutCachedTop =195
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =480
                    UseTheme =0
                    Gradient =0
                    BackColor =14461583
                    BorderWidth =1
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
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =10485
                    Top =135
                    Width =330
                    Height =330
                    TabIndex =3
                    ForeColor =4210752
                    Name ="cmdCloseBd"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Fermeture la base."
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000082c2ea0982c2ea4b82c2ea90 ,
                        0x82c2eade00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000082c2ea2182c2ea7582c2eab782c2eaf982c2eaff82c2eaff ,
                        0x82c2eaff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffa500000000000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffc000000000b17d4a90b17d4affb17d4af0b17d4a36 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffedffffff30b17d4a87b17d4affb17d4af0b17d4a3600000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaffffffffffd7ecf8ff82c2eaff ,
                        0x82c2eaffffffff30b17d4a81b17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff82c2eaff82c2eaff82c2eaffdceef9ffc4e2f5ff82c2eaff ,
                        0x82c2eaffffffff27b17d4a7eb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffe4ffffff27b17d4a84b17d4affb17d4af0b17d4a3900000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffbd00000000b17d4a8db17d4affb17d4af0b17d4a39 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffa500000000000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2ea2182c2ea6f82c2eab782c2eaf982c2eaff82c2eaff ,
                        0x82c2eaff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000082c2ea0982c2ea4e82c2ea96 ,
                        0x82c2eae400000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10485
                    LayoutCachedTop =135
                    LayoutCachedWidth =10815
                    LayoutCachedHeight =465
                    UseTheme =0
                    Gradient =0
                    BackColor =6567968
                    BackTint =100.0
                    BackShade =50.0
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
                Begin Image
                    Left =13495
                    Top =113
                    Width =397
                    Height =397
                    BorderColor =10921638
                    Name ="img_InfoBase"
                    Picture ="ic_BdD.png"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x89504e470d0a1a0a0000000d4948445200000022000000200802000000f8ed3d ,
                        0x9e0000000467414d410000b18f0bfc6105000000097048597300000ec300000e ,
                        0xc301c76fa864000002bf49444154484bcd963f6813511cc793bb4b72f977b9a6 ,
                        0x354d5a51db413a56a14317a98b204e8283a04341040705172711870e0e8208ce ,
                        0x82836eee828b4507958a74d2e02045aca5b1d8cbffcbfdabdfcbefe592d023b9 ,
                        0xbb18f0c3f1f8bddffbf37d77eff77befc2a74f2d86c68f579915513e2a440b7c ,
                        0x44e622f130372d445943875d436b1e588aa5ef98fa4f435b5715d6d06688cc52 ,
                        0x34bd129716a2c9b2697c3754c5323045dd3237b42aebd1013d931c8fa5c89c30 ,
                        0x2f88195e286af5f566857a0e92b994983a1b97dfb72aaf1bfb254b675e6fe4b8 ,
                        0xc8b9c4c4724c7ad3545e36f638e676e3626aea49e5d7f35ac9af06c0100cc470 ,
                        0x4c82ea2019704b9ac13b6169acee190cc1400ca7eaa08ff622b7f058d91e716f ,
                        0x6ecbb3574ac52132e84136451aa628f0d10191b6636ab41427d268125be64e66 ,
                        0x76319626ef50365bd587e56d56f100c9d87be35d03f8eaece01e02ea81b5a5ab ,
                        0x786030d768b8cb20d8efee6fe17955ffc35ca3e12e837c24e3ad5a2623184e26 ,
                        0xb8c814b506928b7ac0c09e93df179437f7278e53d545e6abd640b9144b5f8867 ,
                        0x617cd19a6d7717a408b37ab04f3f51be9acadd94661e65e71f4ccecd4562cfaa ,
                        0xbbd46a0734628e2a007b7eedf7371888729414bb4f8f9c14c3dd0521344e4444 ,
                        0x56e9303c6f7a659cb4b8271f43b9a6fc4089052e8b92dddcc6c9592f74f3a617 ,
                        0x1c0f64408034c04735c8f6f4d227a398061d56f4a1f1d036c089a6769780f4c9 ,
                        0x6cb66a64dcc814ae4b793cabe969f2e0d62123187d32efda7321c09c0d977981 ,
                        0x5ee8536705be70f2a62f04104228710c6376f2007c2e040f0c27babc840004ce ,
                        0x88195cbe98ca25d2bc804b28c87de3eb22407ea00c72df50fd30d4835582e29e ,
                        0x3763e2ff9071223218ce70be50c8937598c9b07039954b84b93d53affbbc4621 ,
                        0x703e9e5d4de73fa895cf5a6d500800fbbc892503ff406db4ea147243641ce807 ,
                        0x6a5cbfeaff8650e82f42c89673332df0fa0000000049454e44ae426082
                    End

                    LayoutCachedLeft =13495
                    LayoutCachedTop =113
                    LayoutCachedWidth =13892
                    LayoutCachedHeight =510
                    TabIndex =26
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =10998
                    Top =170
                    Width =1425
                    Height =750
                    TabIndex =4
                    ForeColor =4210752
                    Name ="cmbLanceCreation"
                    Caption ="Lance la création"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Création du formulaire..."
                    Picture ="ic_Form3.png"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x89504e470d0a1a0a0000000d494844520000001400000010080600000016185f ,
                        0x1b0000053f69545874584d4c3a636f6d2e61646f62652e786d7000000000003c ,
                        0x3f787061636b657420626567696e3d22efbbbf222069643d2257354d304d7043 ,
                        0x656869487a7265537a4e54637a6b633964223f3e0a3c783a786d706d65746120 ,
                        0x786d6c6e733a783d2261646f62653a6e733a6d6574612f2220783a786d70746b ,
                        0x3d22584d5020436f726520352e352e30223e0a203c7264663a52444620786d6c ,
                        0x6e733a7264663d22687474703a2f2f7777772e77332e6f72672f313939392f30 ,
                        0x322f32322d7264662d73796e7461782d6e7323223e0a20203c7264663a446573 ,
                        0x6372697074696f6e207264663a61626f75743d22220a20202020786d6c6e733a ,
                        0x657869663d22687474703a2f2f6e732e61646f62652e636f6d2f657869662f31 ,
                        0x2e302f220a20202020786d6c6e733a70686f746f73686f703d22687474703a2f ,
                        0x2f6e732e61646f62652e636f6d2f70686f746f73686f702f312e302f220a2020 ,
                        0x2020786d6c6e733a746966663d22687474703a2f2f6e732e61646f62652e636f ,
                        0x6d2f746966662f312e302f220a20202020786d6c6e733a786d703d2268747470 ,
                        0x3a2f2f6e732e61646f62652e636f6d2f7861702f312e302f220a20202020786d ,
                        0x6c6e733a786d704d4d3d22687474703a2f2f6e732e61646f62652e636f6d2f78 ,
                        0x61702f312e302f6d6d2f220a20202020786d6c6e733a73744576743d22687474 ,
                        0x703a2f2f6e732e61646f62652e636f6d2f7861702f312e302f73547970652f52 ,
                        0x65736f757263654576656e7423220a202020657869663a436f6c6f7253706163 ,
                        0x653d2231220a202020657869663a506978656c5844696d656e73696f6e3d2232 ,
                        0x30220a202020657869663a506978656c5944696d656e73696f6e3d223136220a ,
                        0x20202070686f746f73686f703a436f6c6f724d6f64653d2233220a2020207068 ,
                        0x6f746f73686f703a49434350726f66696c653d22735247422049454336313936 ,
                        0x362d322e31220a202020746966663a496d6167654c656e6774683d223136220a ,
                        0x202020746966663a496d61676557696474683d223230220a202020746966663a ,
                        0x5265736f6c7574696f6e556e69743d2232220a202020746966663a585265736f ,
                        0x6c7574696f6e3d2239362f31220a202020746966663a595265736f6c7574696f ,
                        0x6e3d2239362f31220a202020786d703a4d65746164617461446174653d223230 ,
                        0x32322d30352d30375430373a34313a30322b30323a3030220a202020786d703a ,
                        0x4d6f64696679446174653d22323032322d30352d30375430373a34313a30322b ,
                        0x30323a3030223e0a2020203c786d704d4d3a486973746f72793e0a202020203c ,
                        0x7264663a5365713e0a20202020203c7264663a6c690a202020202020786d704d ,
                        0x4d3a616374696f6e3d2270726f6475636564220a202020202020786d704d4d3a ,
                        0x736f6674776172654167656e743d22416666696e6974792044657369676e6572 ,
                        0x20312e31302e35220a202020202020786d704d4d3a7768656e3d22323032322d ,
                        0x30352d30375430373a33383a30342b30323a3030222f3e0a20202020203c7264 ,
                        0x663a6c690a20202020202073744576743a616374696f6e3d2270726f64756365 ,
                        0x64220a20202020202073744576743a736f6674776172654167656e743d224166 ,
                        0x66696e6974792050686f746f20312e31302e35220a2020202020207374457674 ,
                        0x3a7768656e3d22323032322d30352d30375430373a34313a30322b30323a3030 ,
                        0x222f3e0a202020203c2f7264663a5365713e0a2020203c2f786d704d4d3a4869 ,
                        0x73746f72793e0a20203c2f7264663a4465736372697074696f6e3e0a203c2f72 ,
                        0x64663a5244463e0a3c2f783a786d706d6574613e0a3c3f787061636b65742065 ,
                        0x6e643d2272223f3e87283e960000018169434350735247422049454336313936 ,
                        0x362d322e31000028917591cf2b445114c73f068d18512c2c2c5e1a6c86fca889 ,
                        0x8dc54c7e15166394c166e69937a3e6c7ebbd3769b255b68a121bbf16fc056c95 ,
                        0xb552444a96b22636e839cf5323997b3be77ceef7de73baf75cf044336ad6acea ,
                        0x816cce3222a321653636a7781fa9c62bd6495b5c35f5c9e9912865c7db0d154e ,
                        0xbcea726a953ff7efa85b4c9a2a54d4080fa9ba61098f094f2c5bbac39bc2cd6a ,
                        0x3abe287c2c1c30e482c2d78e9e70f9c9e194cb1f0e1bd148183c8dc24aea1727 ,
                        0x7eb19a36b2c2f272fcd94c41fdb98ff3125f3237332db14dac159308a3845018 ,
                        0x679830417a19141fa48b3eba654599fc9eeffc29f292ab8ad72962b0448a3416 ,
                        0x01510b523d2951133d293343d1e9ffdfbe9a5a7f9f5bdd1782ea07db7e6907ef ,
                        0x067caedbf6fbbe6d7f1e40e53d9ce54af9f93d1878157dbda4f977a161154ece ,
                        0x4b5a620b4ed7a0e54e8f1bf16fa952cca369f07c04f53168ba84da79b7673ffb ,
                        0x1cde427445beea02b677a043ce372c7c0159c967e0fb3f7a7900000009704859 ,
                        0x7300000ec400000ec401952b0e1b0000013149444154388d63343131f9afa0a0 ,
                        0xc0400df0e0c103061605050586868686420606860b149a67d0d0d0d0cf02e55c ,
                        0xd0d1d13940896957ae5c6160606060604113ec6760603020d34c010c039357de ,
                        0x5af875dff485578eee26d9fb57ae5c71606060d8cf842cc82ca6dccf6911b9df ,
                        0x20288b5c57a2ba90919d9b814ddd4e805950fabcb9b68be1c9e6a00b3366cc10 ,
                        0x60c01f0c1f323232e03e62c2aa848995e1c7f98dfd14bbf0ffafef0c7f9e5efd ,
                        0xf0e7c52dc78b5b165e60606060c8c8c8f8c0c0c070802c03ffbd7b5cf8edd05c ,
                        0x06722205ab8109628f18186283051862831da042171820e1d7cfc0c0f0019721 ,
                        0x1919198e580dc4030a3332320e6093983163c67e9c2e448e2d240d1f181818fa ,
                        0x67cc98518fc332143d045d08b5c491903a74030d6079910260c0c0c0c0c048ed ,
                        0xe20b00bd25647f2f633f710000000049454e44ae426082
                    End

                    LayoutCachedLeft =10998
                    LayoutCachedTop =170
                    LayoutCachedWidth =12423
                    LayoutCachedHeight =920
                    PictureCaptionArrangement =4
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
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =10095
                    Top =150
                    Width =330
                    Height =330
                    TabIndex =5
                    ForeColor =4210752
                    Name ="cmbOuvreBase"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Ouvrir la base"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000002114881021148870211488bf ,
                        0x211488ef211488ff211488ff211488ff211488ff211488ff211488ef211488bf ,
                        0x2114887021148810000000000000000000000000211488ef211488ff211488ff ,
                        0x211488ff211488ff211488ff211488ff211488ff211488ff211488ff211488ff ,
                        0x211488ff211488ef000000000000000000000000100a44ff100a44ff100a44ff ,
                        0x100a44ff100a44ff100a44ff100a44ff190f66ff211488ff211488ff211488ff ,
                        0x211488ff211488ff3120afef3120afff3120afff3120afff3120afff3120afff ,
                        0x3120afff3120afff3120afff2f1fa8ff100a44ff211488ff211488ff211488ff ,
                        0x211488ff211488ff3120afff3120afff3120afff3120afff3120afff3120afff ,
                        0x3120afff3120afff3120afff3120afff181057ff3120afff301fadff2d1da5ff ,
                        0x281999ff22158aff3120afff3120afffe5e3f5ffbfb9e6ff3120afff3120afff ,
                        0xbfb9e6ffe5e3f5ff3120afff3120afff181057ff3120afff3120afff3120afff ,
                        0x3120afff301fadff3120afff3120afff9890d7ffffffffffccc7ebffccc7ebff ,
                        0xffffffff9890d7ff3120afff3120afff181057ff3120afff3120afff3120afff ,
                        0x3120afff3120afff3120afff3120afff4b3cb9ffffffffffccc7ebffccc7ebff ,
                        0xffffffff4b3cb9ff3120afff3120afff181057ff3120afff3120afff3120afff ,
                        0x3120afff3120afff3120afff3120afff3120afffccc7ebffbfb9e6ffbfb9e6ff ,
                        0xd8d5f0ff3120afff3120afff3120afff302764ff604fc9ff5d4cc7ff5443c3ff ,
                        0x4635baff3423b1ff3120afff3120afff3120afff7e74cdffffffffffffffffff ,
                        0x7e74cdff3120afff3120afff3120afff302764ff604fc9ff604fc9ff604fc9ff ,
                        0x604fc9ff5d4cc7ff3120afff3120afff3120afff3e2eb4ffffffffffffffffff ,
                        0x3e2eb4ff3120afff3120afff3120afff302764ff604fc9ff604fc9ff604fc9ff ,
                        0x604fc9ff604fc9ff3120afff3120afff3120afff3120afff3120afff3120afff ,
                        0x3120afff3120afff3120afff3120afff332a6aff604fc9ff604fc9ff604fc9ff ,
                        0x604fc9ff604fc9ff3120afef3120afff3120afff3120afff3120afff3120afff ,
                        0x3120afff3120afff3120afff3726b2ff9580e0ff9580e0ff927ddfff8874daff ,
                        0x7764d3ff6352caff000000000000000000000000927ddfff9580e0ff9580e0ff ,
                        0x9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff ,
                        0x9580e0ff927ddfff0000000000000000000000009580e0ef9580e0ff9580e0ff ,
                        0x9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff ,
                        0x9580e0ff9580e0ef0000000000000000000000009580e0109580e0709580e0bf ,
                        0x9580e0ef9580e0ff9580e0ff9580e0ff9580e0ff9580e0ff9580e0ef9580e0bf ,
                        0x9580e0709580e010
                    End

                    LayoutCachedLeft =10095
                    LayoutCachedTop =150
                    LayoutCachedWidth =10425
                    LayoutCachedHeight =480
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
                Begin OptionGroup
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =10197
                    Top =9242
                    Width =954
                    Height =433
                    TabIndex =21
                    BorderColor =10921638
                    Name ="BoxOptImages"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="2"
                    GridlineColor =10921638

                    LayoutCachedLeft =10197
                    LayoutCachedTop =9242
                    LayoutCachedWidth =11151
                    LayoutCachedHeight =9675
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =7427
                            Top =9355
                            Width =2700
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_BoxOptImages"
                            Caption ="Image CommandButton :"
                            GridlineColor =10921638
                            LayoutCachedLeft =7427
                            LayoutCachedTop =9355
                            LayoutCachedWidth =10127
                            LayoutCachedHeight =9670
                        End
                        Begin ToggleButton
                            OverlapFlags =255
                            Left =10254
                            Top =9298
                            Width =405
                            Height =345
                            OptionValue =1
                            ForeColor =4210752
                            Name ="optImgOn"
                            Caption ="On"
                            FontName ="Calibri"
                            GridlineColor =10921638

                            LayoutCachedLeft =10254
                            LayoutCachedTop =9298
                            LayoutCachedWidth =10659
                            LayoutCachedHeight =9643
                            UseTheme =0
                            BackColor =14461583
                            BorderColor =14461583
                            HoverColor =15189940
                            PressedColor =9917743
                            HoverForeColor =4210752
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =247
                            Left =10650
                            Top =9298
                            Width =420
                            Height =345
                            TabIndex =1
                            OptionValue =2
                            ForeColor =4210752
                            Name ="optImgOff"
                            Caption ="Off"
                            FontName ="Calibri"
                            GridlineColor =10921638

                            LayoutCachedLeft =10650
                            LayoutCachedTop =9298
                            LayoutCachedWidth =11070
                            LayoutCachedHeight =9643
                            UseTheme =0
                            BackColor =14461583
                            BorderColor =14461583
                            HoverColor =15189940
                            PressedColor =9917743
                            HoverForeColor =4210752
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                            Overlaps =1
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =6576
                    Top =9638
                    Width =230
                    Height =1417
                    BackColor =15921906
                    BorderColor =10921638
                    Name ="boxCacheCtrImgCmb"
                    GridlineColor =10921638
                    LayoutCachedLeft =6576
                    LayoutCachedTop =9638
                    LayoutCachedWidth =6806
                    LayoutCachedHeight =11055
                    BackShade =95.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Left =286
                    Top =169
                    Width =2565
                    Height =315
                    FontSize =12
                    BackColor =5855577
                    BorderColor =8355711
                    Name ="lbl_txtBdd"
                    Caption ="Base de données :"
                    FontName ="Arial"
                    GridlineColor =10921638
                    LayoutCachedLeft =286
                    LayoutCachedTop =169
                    LayoutCachedWidth =2851
                    LayoutCachedHeight =484
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =4365
                    Top =170
                    Width =3561
                    Height =315
                    ForeColor =4210752
                    Name ="lbl_InfoBaseNonSelect"
                    Caption ="Sélectionnez une base de données..."
                    GridlineColor =10921638
                    LayoutCachedLeft =4365
                    LayoutCachedTop =170
                    LayoutCachedWidth =7926
                    LayoutCachedHeight =485
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeTint =75.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2948
                    Top =566
                    Width =7881
                    Height =315
                    TabIndex =1
                    BackColor =15921906
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtBddSauve"
                    GridlineColor =10921638

                    LayoutCachedLeft =2948
                    LayoutCachedTop =566
                    LayoutCachedWidth =10829
                    LayoutCachedHeight =881
                    BackShade =95.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =283
                            Top =566
                            Width =2550
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="lbl_txtBddSauve"
                            Caption ="Sauvegarde :"
                            GridlineColor =10921638
                            LayoutCachedLeft =283
                            LayoutCachedTop =566
                            LayoutCachedWidth =2833
                            LayoutCachedHeight =881
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13375
                    Top =5045
                    Width =231
                    Height =315
                    TabIndex =25
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtCacheFields"
                    GridlineColor =10921638

                    LayoutCachedLeft =13375
                    LayoutCachedTop =5045
                    LayoutCachedWidth =13606
                    LayoutCachedHeight =5360
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="F_Pied"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------
' Name:     Form_F_CreateForm
' Kind:     Document VBA
' Purpose:  Formulaire de définition des options pour la création d'un formulaire pour la classe CSordFormColumn
' Author:   Laurent
' Date:     28/04/2022
' DateMod:  15/05/2022-12:12
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//----------------------------------       VAR       ------------------------------

    Private Const LBL_COLOR As Long = 6710886       '// ForeColor label, 'Texte 1, Plus clair 40%.

    Private m_cCreate       As CCreateFormContinu
    Private m_cUtil         As CUtilitaires

    Private m_bErrSaisie    As Boolean  '// Indique erreur de saisie.

    '// Rst pour affichage des infos sur le ctr sélectionné.
    Private Const TAB_INFO  As String = "T_Info"
    Private m_oRst          As DAO.Recordset
    Private Const ID_INF    As String = "[ID_Info]='"
    Private Const C_COULDEF As Long = 16777215      '// BackColor label sélectionné.
    Private Const C_COULSEL As Long = 14610923      '// BackColor label déselectionné.
    Private m_sCtrPrec      As String

'//---------------------------------------------------------------------------------------

'//==================================       EVENT       ==================================
Private Sub Form_Load()

    '// Initialisation de la classe.
    Set m_cCreate = New CCreateFormContinu
    Set m_cUtil = New CUtilitaires

    '// Table info controles.
    Set m_oRst = CurrentDb.OpenRecordset(TAB_INFO, dbOpenSnapshot, dbReadOnly)

    '// Masque la barre de navigation et la barre des boutons.
    'NavigationPane False
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    Me.boxCacheCtrImgCmb.Width = 6395       '// Masque les textBox pour les images.

    '// Applique les valeurs par défaut...
    RazForm

    '// Affiche le message select BdD.
    Me.TimerInterval = 2000

End Sub

Private Sub Form_Close()
    On Error GoTo ERR_Form_Close

    Screen.MousePointer = 11    '// Hourglass.

    If (Not m_oRst Is Nothing) Then
        m_oRst.Close
        Set m_oRst = Nothing
    End If

    '// Déclenche class_Terminate()
    Set m_cCreate = Nothing

    NavigationPane True
    DoCmd.ShowToolbar "Ribbon", acToolbarYes

SORTIE_Form_Close:
    DoCmd.Echo True
    Screen.MousePointer = 0
    Exit Sub

ERR_Form_Close:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.Form_F_CreateForm.Form_Close, ligne " & Erl & "."
    Resume SORTIE_Form_Close
End Sub

Private Sub cmbSelectBdd_Click()
    On Error GoTo ERR_cmbSelectBdd

    Dim bRet     As Boolean
    Dim sBaseSel As String
    Dim vTmp     As Variant  '// Pour Split de sBackup.
    Dim sRep     As String

    '// Met en pause le msg pas de bd.
    Me.TimerInterval = 0
    Me.lbl_InfoBaseNonSelect.Visible = False
    
    '// Séléction de la base à utiliser.
    sBaseSel = OuvreBoite("MS Access", "*.accdb", , CurrentProject.Path, FD_TypeFilePicker)
    If (sBaseSel = vbNullString) Then GoTo SORTIE_cmbSelectBdd

    DoCmd.Echo False
    Screen.MousePointer = 11            '// Hourglass.

    '// Création Access.Application, si pas déjà fait.
    If (m_cCreate.MsAppIsUp = False) Then
        bRet = m_cCreate.OpenMsApp()
        If (bRet = False) Then GoTo SORTIE_cmbSelectBdd
    End If

    If (m_cCreate.MsBaseIsOpen = False) Then
        bRet = m_cCreate.OpenMsBase(sBaseSel) '// Ouverture de la base.
    End If

    If (bRet = False) Then
        '// Problème détecter, on ferme tout, RaZ et on sort.
        m_cCreate.CloseMsBase True
        RazForm
        GoTo SORTIE_cmbSelectBdd
    End If

    RazForm True                        '// Reset les valeurs...

    Me.txtBdd.Enabled = True
    txtBdd = sBaseSel
    Me.txtBdd.SetFocus
    Me.cmbOuvreBase.Visible = False
    MaJlisteObjets                      '// Rempli la liste des objets ...
    
    '// Détermine le nonm du fichier de la prochaine sauvegarde...
    sRep = GetBackupFileName(sBaseSel)
    '// NOTE retourne folder;backup;base
    vTmp = Split(sRep, ";")
    Me.txtBddSauve = vTmp(0) & vTmp(1) '// folder + backup.

SORTIE_cmbSelectBdd:
    Screen.MousePointer = 0
    DoCmd.Echo True
    Exit Sub

ERR_cmbSelectBdd:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TriSurFormContinu.Form_F_CreateForm.cmbSelectBdd_Click, ligne " & Erl & "."
    Resume SORTIE_cmbSelectBdd
End Sub

Private Sub cmdCloseBd_Click()

    Screen.MousePointer = 11    '// Hourglass.
    '// Ferme la base en cours, réinitialise les champs par défaut...
    m_cCreate.CloseMsBase
    RazForm
    Screen.MousePointer = 0
    Me.txtBdd = "Sélectionnez un base..."

End Sub

Private Sub cmbOuvreBase_Click()
    On Error GoTo ERR_cmbOuvreBase_Click

    Dim dShell As Double
    dShell = Shell("MSAccess.exe " & Me.txtBdd, 3)
    
SORTIE_cmbOuvreBase_Click:
    Exit Sub

ERR_cmbOuvreBase_Click:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  CSord.Form_F_CreateForm.cmbOuvreBase_Click, ligne " & Erl & "."
    Resume SORTIE_cmbOuvreBase_Click
End Sub

'// MàJ de la liste des champs de l'objet sélectionné.
Private Sub lstObjets_AfterUpdate()

    If (Me.lstObjets = Null) Then Exit Sub

    If ((m_cCreate.MsAppIsUp = False) Or (m_cCreate.MsBaseIsOpen = False)) Then Exit Sub

    Dim sListeVal As String

    '// Rempli(liste de valeurs) la liste lstFields...
    sListeVal = ObjectFieldsToListVal(Me.lstObjets, Me.lstObjets.Column(0), m_cCreate.objMsBase)
    Me.lstFields.RowSource = sListeVal

    Me.txtFormName = "F_" & Me.lstObjets
    Me.txtFormSource = Me.lstObjets
    Me.lstObjets.Controls(0).ForeColor = LBL_COLOR
    Me.txtFormName.Controls(0).ForeColor = LBL_COLOR
    Me.txtFormSource.Controls(0).ForeColor = LBL_COLOR
    Me.lbl_lstFieldsInfo.Caption = "Champs de : " & Me.lstObjets

End Sub

Private Sub BoxOptImages_AfterUpdate()
    Select Case BoxOptImages
        Case 1
            Me.boxCacheCtrImgCmb.Visible = False
            Me.txtPicFolder.TabStop = True
            Me.txtPicAsc.TabStop = True
            Me.txtPicDesc.TabStop = True
            Me.txtPicFolder.SetFocus
        Case 2
            Me.boxCacheCtrImgCmb.Visible = True
            Me.txtPicFolder.TabStop = False
            Me.txtPicFolder = vbNullString
            Me.txtPicAsc.TabStop = False
            Me.txtPicAsc = vbNullString
            Me.txtPicDesc.TabStop = False
            Me.txtPicDesc = vbNullString
    End Select
End Sub

Private Sub cmbSelectPicFolder_Click()
    '// Sélection du dossier img pour commandButton.
    Dim sRet As String

    sRet = OuvreBoite("Dossier des images...", , , , FD_TypeFolderPicker)
    If (sRet = vbNullString) Then Exit Sub

    Me.txtPicFolder = sRet & "\"

End Sub

Private Sub cmbSelectPicAsc_Click()
    '// Sélection de l'image ASC.
    Dim sRet As String
    Dim lTmp As Long
    Dim sTmp As String

    sRet = OuvreBoite("Image...", "*.png,*.jpg,*.bmp", "Sélectionnez l'image ASC", Me.txtPicFolder, FD_TypeFilePicker)
    If (sRet = vbNullString) Then Exit Sub
    
    '// Récupère que le fichier.
    lTmp = Len(sRet) - InStrRev(sRet, "\")
    sTmp = Right$(sRet, lTmp)

    Me.txtPicAsc = sTmp
End Sub

Private Sub cmbSelectPicDesc_Click()
    '// Sélection de l'image DESC.
    Dim sRet As String
    Dim lTmp As Long
    Dim sTmp As String

    sRet = OuvreBoite("Image...", "*.png,*.jpg,*.bmp", "Sélectionnez l'image DESC", Me.txtPicFolder, FD_TypeFilePicker)
    If (sRet = vbNullString) Then Exit Sub

    '// Récupère que le fichier.
    lTmp = Len(sRet) - InStrRev(sRet, "\")
    sTmp = Right$(sRet, lTmp)

    Me.txtPicDesc = sTmp
End Sub

' ----------------------------------------------------------------
' Procedure Nom:    cmbLanceCreation_Click
' Sujet:            Lance la création du formulaire.
' Procedure Kind:   Sub
' Procedure Access: Private
' Références:       Lance la création du formulaire.
'
'=== Paramètres ===
'==================
'
'
' Author:  Laurent
' Date:    03/05/2022 - 15:35
' DateMod:
'
' !Use! :
' ----------------------------------------------------------------
Private Sub cmbLanceCreation_Click()

    Dim bRet    As Boolean
    Dim sArg    As String

    bRet = VerifSaisie()                                        '// Contrôle des saisies...
    If (bRet = True) Then Exit Sub

    Me.txtCacheFields = ExtraireFieldsList(True)
    DoCmd.OpenForm "F_Info", acNormal, , , , acDialog, "RES"    '// Affiche résumé de qu'il vas être fait.

    If Me.txtCacheFields = "ANNULER" Then Exit Sub

    Screen.MousePointer = 11    '// Hourglass.
    DoCmd.Echo False

    bRet = m_cCreate.CloseMsBase                                '// Fermeture de la base pour sauvegarde...
    If (bRet = False) Then GoTo SORTIE_LanceCreation

    bRet = CopyFile(m_cCreate.GetBaseFullName, Me.txtBddSauve)  '// Sauvegarde sous le nom txtBddSauve.
    If (bRet = False) Then GoTo SORTIE_LanceCreation

    bRet = m_cCreate.OpenMsBase(m_cCreate.GetBaseFullName)      '// Réouverture de la base...
    If (bRet = False) Then GoTo SORTIE_LanceCreation

    bRet = ActualiseOptionsClass()                              '// MàJ des propriétés de la classe...
    If (bRet = False) Then Exit Sub

    bRet = m_cCreate.LanceCreation()                            '// Lance la création...

    Me.txtBdd.SetFocus
    Me.cmbLanceCreation.Enabled = False
    
    m_cCreate.CloseMsBase True                                  '// On ferme tout, base et application...

    RazForm                                                     '// RaZ des valeurs du form...

SORTIE_LanceCreation:
    Screen.MousePointer = 0
    DoCmd.Echo True

    If bRet Then
        MsgBox "Opération terminée avec succès", vbInformation, "Création du Formulaire"
        '// Permet d'ouvrir la base avec le bouton cmbOuvreBase.
        Me.cmbOuvreBase.Visible = True
        Me.txtBdd = m_cCreate.GetBaseFullName
    Else
        Me.txtBdd = vbNullString
    End If

End Sub

'//=======================================================================================

'// ################################ PRIVATE SUB/FUNC ####################################

'// Rempli la liste des objets, avec ceux trouver dans la BdD.
Private Sub MaJlisteObjets()

    '// Rempli la liste des objets de la base (Tables-Requêtes)
    Dim sListVal As String
    sListVal = ListObjects(Tables_Local, True, Tables_Linked, QueriesType, m_cCreate.objMsBase)

    Me.lstObjets.RowSource = sListVal
    Me.lstObjets = vbNullString

End Sub

'// Applique les valeurs par défaut.
Private Sub RazForm(Optional bActive As Boolean)

    Me.txtBddSauve = vbNullString
    Me.txtFormName = "F_"
    Me.txtClasseName = m_cCreate.OptVarClasse
    Me.txtFunctionName = m_cCreate.OptFunctionName

    Me.txtTbSuffix = m_cCreate.OptTextBoxSuffix
    Me.txtCmbSuffix = m_cCreate.OptCmbSuffix
    
    Me.BoxOptImages = 2
    BoxOptImages_AfterUpdate        '// Masque les txt images.

    Me.lstObjets.RowSource = vbNullString
    Me.lstObjets.SetFocus
    Me.lstFields.RowSource = vbNullString
    Me.lbl_lstFieldsInfo.Caption = "Champs"

    Me.cmdCloseBd.Visible = bActive
    Me.cmbSelectBdd.Visible = Not bActive
    Me.cmbLanceCreation.Enabled = bActive
    Me.cmbSelectPicFolder.Enabled = bActive
    Me.cmbSelectPicAsc.Enabled = bActive
    Me.cmbSelectPicDesc.Enabled = bActive
    Me.BoxOptImages.Enabled = bActive

End Sub

'// Stock les options pour la création du formulaire.
Private Function ActualiseOptionsClass() As Boolean

    Dim sVerif  As String
    Dim vItem   As Variant

    sVerif = Nz(Me.lstObjets, vbNullString)
    
    With m_cCreate
        .OptFormName = Me.txtFormName
        .OptFormSource = Me.lstObjets
        .OptVarClasse = Me.txtClasseName
        .OptFunctionName = Me.txtFunctionName
        .OptTextBoxPrefix = Nz(Me.txtTbPrefix, vbNullString)
        .OptTextBoxSuffix = Nz(Me.txtTbSuffix, vbNullString)
        .OptCmbPrefix = Nz(Me.txtCmbPrefix, vbNullString)
        .OptCmbSuffix = Nz(Me.txtCmbSuffix, vbNullString)
        .OptPictureFolder = Nz(Me.txtPicFolder, vbNullString)
        .OptPictureAsc = Nz(Me.txtPicAsc, vbNullString)
        .OptPictureDesc = Nz(Me.txtPicDesc, vbNullString)
    End With

    '// Stock les champs sélectionés de la liste lstFields.
    ExtraireFieldsList

    ActualiseOptionsClass = True

End Function

'// Affiche des infos contenu dans la table 'T_Info', suivant le controle en cours.
'//
'// si sID est indiquer on utilise pas ActiveControl.Name, mais la valeur de sID.
'//
'// AfficheInfo est défini sur OnGotFocus du controle avec =AfficheInfo("")
'//
Private Function AfficheInfo(Optional sID As String = vbNullString)

    If m_bErrSaisie Then Exit Function      '// Erreur de saisie, laisse les infos erreur afficher.

    '// Restaure le ctr précedent, applique la backColor.
    If (m_sCtrPrec <> vbNullString) Then Me(m_sCtrPrec).BackColor = C_COULDEF
    m_sCtrPrec = Me.ActiveControl.Name
    Me(m_sCtrPrec).BackColor = C_COULSEL

    Dim sFind As String

    '// Pour les préfixe/suffixe on affiche les mêmes infos (TextBox et CommandButton)
    sFind = IIf(sID <> vbNullString, ID_INF & sID & "'", ID_INF & m_sCtrPrec & "'")

    m_oRst.FindFirst sFind
    If (m_oRst.NoMatch) Then Exit Function

    txtInfoTitre = m_oRst.Fields("InfoTitre")
    txtInfoTxt = m_oRst.Fields("InfoTexte")

End Function

' ----------------------------------------------------------------
' Procedure Nom:    VerifSaisie
' Sujet:            Controle de la validité des saisies du formulaire.
' Procedure Kind:   Sub
' Procedure Access: Private
'
' Author:  Laurent
' Date:    03/05/2022 - 08:44
' DateMod: 15/05/2022-14:18
'
' ----------------------------------------------------------------
Private Function VerifSaisie() As Boolean

    Dim oCtr    As Control
    Dim sName   As String   '// Nom du control.
    Dim bErrChk As Boolean  '// True si erreur saisie

    m_bErrSaisie = False

    '// Parcour les ctr du form.
    For Each oCtr In Me.Controls

        sName = oCtr.Name

        '// Controle la saisie des ListBox.
        If (oCtr.ControlType = acListBox) Then

            oCtr.Controls(0).ForeColor = LBL_COLOR
            '// Max 10 CommandButton de créer, sinon form trop petit.
            If ((oCtr.ItemsSelected.Count = 0) Or (oCtr.ItemsSelected.Count > 11)) Then bErrChk = True

        End If

        '// Controle la saisie des TextBox
        If (oCtr.ControlType = acTextBox) Then

            If (oCtr <> vbNullString) Then sName = "*"  '// Remet le label avec la couleur normal.

                Select Case sName
                    Case "*"            '// Ignore le controle.
                        If (oCtr.Controls.Count > 0) Then oCtr.Controls(0).ForeColor = LBL_COLOR

                    Case Me.txtFormName.Name, Me.txtClasseName.Name, Me.txtFunctionName.Name
                        bErrChk = True

                    Case Me.txtTbPrefix.Name, Me.txtTbSuffix.Name
                        If (IsNull(Me.txtTbPrefix) And IsNull(Me.txtTbSuffix)) Then
                            bErrChk = True
                        Else
                            oCtr.Controls(0).ForeColor = LBL_COLOR
                        End If

                    Case Me.txtCmbPrefix.Name, Me.txtCmbSuffix.Name
                        If (IsNull(Me.txtCmbPrefix) And IsNull(Me.txtCmbSuffix)) Then
                            bErrChk = True
                        Else
                            oCtr.Controls(0).ForeColor = LBL_COLOR
                        End If

                    Case Me.txtPicFolder.Name, Me.txtPicAsc.Name, Me.txtPicDesc.Name
                        If (Me.BoxOptImages = 1) Then bErrChk = True

                End Select
        End If

        If bErrChk Then
            oCtr.Controls(0).ForeColor = 2366701   '// label du control en rouge.
            m_bErrSaisie = True
        End If

        bErrChk = False     '/// loop.
    Next
    Set oCtr = Nothing

    '// Affiche info erreur de saisie pendant 10 sec.
    If m_bErrSaisie Then
        Me.txtInfoTitre.BorderStyle = 1
        Me.txtInfoTitre.BorderColor = 2366701   '// Rouge
        Me.txtInfoTitre = "Erreurs de saisie :"
        Me.txtInfoTxt = "Les champs indiquer en rouge doivent obligatoirement contenir une valeur."
        Me.TimerInterval = 10000                '// Affiche l'erreur 10 secondes.
    Else
        Me.txtInfoTitre.BorderColor = 10921638  '// Arrière-plan 1, Plus sombre 35%.
        Me.txtInfoTitre.BorderStyle = 0
        Me.txtInfoTitre = vbNullString
        Me.txtInfoTxt = vbNullString
    End If

    VerifSaisie = m_bErrSaisie

End Function

'// Affiche le résumé des opérations.
Private Function ExtraireFieldsList(Optional bRetourList As Boolean = False) As String

    Dim sFld        As String
    Dim sFields()   As String
    Dim lInd        As Long
    Dim lFor        As Long
    Dim vItem       As Variant

    '// Stock les champs sélectionés de la liste lstFields.
    For Each vItem In lstFields.ItemsSelected
        m_cCreate.AddField = lstFields.ItemData(vItem)
    Next vItem

    '// Retourne la liste des champs sélectionnés.
    If bRetourList Then
        sFields = m_cCreate.GetFields
        lInd = UBound(sFields)
        For lFor = 0 To lInd
            sFld = sFld & sFields(lFor) & ", "
        Next lFor
        ExtraireFieldsList = sFld
    End If

End Function

'// Affiche le message aucune bd select(2sec), ou Efface le message d'erreur au bout de 10 sec.
Private Function Raz_bErr()

    '// Msg pas de Bd ouverte, ce msg est afficher qu'à l'ouverture du form.
    If (IsNull(Me.txtBdd)) Then
        lbl_InfoBaseNonSelect.Visible = Not Me.lbl_InfoBaseNonSelect.Visible
        Exit Function
    End If

    '// Efface msg d'erreur.
    m_bErrSaisie = False
    Me.txtInfoTxt = vbNullString
    Me.txtInfoTitre = vbNullString
    Me.txtInfoTitre.BorderColor = 10921638
    Me.txtInfoTitre.BorderStyle = 0
    Me.TimerInterval = 0

End Function

Private Function RestaureLabelTxt()
    With Me.ActiveControl
        If Not IsNull(.Value) Then
            .Controls(0).ForeColor = LBL_COLOR
        End If
    End With
End Function

Private Function RestaureLabelLst()
    With Me.ActiveControl
        If ((.ItemsSelected.Count > 0) Or (.ItemsSelected.Count < 11)) Then
            .Controls(0).ForeColor = LBL_COLOR
        End If
    End With
End Function
'// ######################################################################################
