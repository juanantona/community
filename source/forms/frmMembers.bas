Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7653
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =4170
    Top =2520
    Right =12360
    Bottom =9660
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x60d7bae5dddfe440
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
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
            BorderColor =16777215
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
            GridlineColor =12632256
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
        Begin FormHeader
            Height =1134
            BackColor =15849926
            Name ="EncabezadoDelFormulario"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =737
                    Top =113
                    Width =6066
                    Height =850
                    FontSize =16
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Etiqueta5"
                    Caption ="COMUNIDAD 11"
                    GridlineColor =10921638
                    LayoutCachedLeft =737
                    LayoutCachedTop =113
                    LayoutCachedWidth =6803
                    LayoutCachedHeight =963
                End
            End
        End
        Begin Section
            Height =4875
            Name ="Detalle"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =1695
                    Top =360
                    Width =5235
                    Height =1695
                    BorderColor =10921638
                    Name ="Secundario0"
                    SourceObject ="Form.sfrmMembersDataView"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1695
                    LayoutCachedTop =360
                    LayoutCachedWidth =6930
                    LayoutCachedHeight =2055
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =360
                            Width =1275
                            Height =1695
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etiqueta1"
                            Caption ="Hermanos:"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =2055
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2267
                    Top =2381
                    Width =1984
                    Height =453
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdShowBrothers"
                    Caption ="Ver grupo"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2267
                    LayoutCachedTop =2381
                    LayoutCachedWidth =4251
                    LayoutCachedHeight =2834
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1656
                    Top =3288
                    Width =5656
                    Height =1125
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtGroups"
                    GridlineColor =10921638

                    LayoutCachedLeft =1656
                    LayoutCachedTop =3288
                    LayoutCachedWidth =7312
                    LayoutCachedHeight =4413
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =3288
                            Width =705
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etiqueta8"
                            Caption ="Grupos"
                            GridlineColor =10921638
                            LayoutCachedLeft =396
                            LayoutCachedTop =3288
                            LayoutCachedWidth =1101
                            LayoutCachedHeight =3603
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            Name ="PieDelFormulario"
            AutoHeight =1
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
Option Compare Database

Private Sub cmdShowBrothers_Click()
  Dim cntCurrentDB As ADODB.Connection
  Dim rsoMembers As ADODB.recordset
  Dim strGroup As String
  
  'Instancia conexión a BD actual
   Set cntCurrentDB = CurrentProject.Connection
   Set rsoMembers = New ADODB.recordset
   rsoMembers.Open "SELECT * FROM lstMembers", cntCurrentDB, adOpenKeyset, adLockOptimistic
   
   Dim strm As ADODB.Stream
   Set strm = New ADODB.Stream
   rsoMembers.Save strm
   Dim rsoCopy As ADODB.recordset
   Set rsoCopy = New ADODB.recordset
   rsoCopy.Open strm
   
   Dim groupNumber As Integer
   Dim groupMember As Integer
   
   groupNumber = 2
   groupMember = 2
   
   Dim i, j, k, rd As Integer
      
   For i = 1 To groupNumber
     For j = 1 To groupMember
       k = 1
       rd = random(rsoCopy.RecordCount)
       rsoCopy.MoveFirst
       Do Until rsoCopy.EOF = True
         If k = rd Then
           strGroup = strGroup & rsoCopy.Fields!FirstNameField & ", "
           rsoCopy.Delete
         End If
         rsoCopy.MoveNext
         k = k + 1
       Loop
       
    Next j
    txtGroups.SetFocus
    strGroup = vbCrLf & strGroup
    txtGroups.Text = txtGroups.Text & strGroup
    strGroup = ""
   Next i
      
   
   
   Set rsoMembers = Nothing
   cntCurrentDB.Close

End Sub

Public Function random(recorCount As Integer) As Integer
  Randomize
  random = Int(recorCount * Rnd) + 1
End Function
