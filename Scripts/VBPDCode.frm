VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{B8F3F837-A5F3-11D5-9768-0050DAC05691}#2.1#0"; "HFCSRecordCountControl.ocx"
Object = "{45655353-8810-11D5-82CE-10005A7C7E5C}#1.3#0"; "HFCSWarningPopup.ocx"
Object = "{4C28E4AB-0CBA-11D5-96F3-0050DAC05691}#33.0#0"; "HFCSBayControl.ocx"
Object = "{A7A71542-28F8-11D5-9DC2-10005A7238BB}#7.0#0"; "HFCSWkEndingDate.ocx"
Object = "{430A1B88-1B5E-4470-8BC2-8F89EFF2ACA9}#6.1#0"; "HFCSTextBoxControl.ocx"
Object = "{E3AFC15F-69D6-4CC7-BFAE-7859ABAC6ED2}#5.10#0"; "HFCSTBManager.ocx"
Begin VB.Form frmPDD
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9270
   ClientLeft      =   450
   ClientTop       =   1485
   ClientWidth     =   14820
   BeginProperty Font
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   14820
   Begin VB.TextBox TextEnd
      Height          =   285
      Left            =   1.41600e5
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CheckBox chkDriverNotified
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEloc
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12600
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTractorNumber
      Enabled         =   0   'False
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10320
      MaxLength       =   7
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDriverName
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtJobNr
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame frmBayData
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5595
      Left            =   240
      TabIndex        =   35
      Top             =   3600
      Width           =   14415
      Begin TrueOleDBGrid70.TDBGrid grdBayDetails
         Height          =   955
         Index           =   0
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   1693
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
        Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
        Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   3
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         AllowArrows     =   0   'False
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin VB.CheckBox chkExtraDolly
         Caption         =   "Extra Dolly"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   10
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CheckBox chkExtraDolly
         Caption         =   "Extra Dolly"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin TrueOleDBGrid70.TDBDropDown drpValTrailerType
         CausesValidation=   0   'False
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   39
         Top             =   5040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   953
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
        Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   1
         BorderStyle     =   1
         ColumnHeaders   =   -1  'True
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   ""
         DataField       =   ""
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   15790320
         ValueTranslate  =   0   'False
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=176,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBDropDown drpValTrailerType
         CausesValidation=   0   'False
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   38
         Top             =   3240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   953
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   1
         BorderStyle     =   1
         ColumnHeaders   =   -1  'True
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   ""
         DataField       =   ""
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   15790320
         ValueTranslate  =   0   'False
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=156,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBDropDown drpValTrailerType
         CausesValidation=   0   'False
         Height          =   495
         Index           =   0
         Left            =   4800
         TabIndex        =   37
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   953
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   1
         BorderStyle     =   1
         ColumnHeaders   =   -1  'True
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   ""
         DataField       =   ""
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   15790320
         ValueTranslate  =   0   'False
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
        _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin HFCSBayControl.BayControl ctlBayNr
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   11
         Top             =   5040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Object.Width           =   605
         GateBay         =   -1  'True
      End
      Begin HFCSBayControl.BayControl ctlBayNr
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   3240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Object.Width           =   605
         GateBay         =   -1  'True
      End
      Begin HFCSBayControl.BayControl ctlBayNr
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   6
         Top             =   1440
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Object.Width           =   605
         GateBay         =   -1  'True
      End
      Begin VB.Frame Frame1
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   135
         TabIndex        =   23
         Top             =   4680
         Width           =   2040
         Begin HFCSTextBoxControl.HFCSTextBox txtDolly
            Height          =   285
            Index           =   2
            Left            =   825
            TabIndex        =   13
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483643
            Enabled         =   -1  'True
            ForeColor       =   -2147483640
            BorderStyle     =   1
            DataField       =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataMember      =   ""
            Locked          =   0   'False
            MaxLength       =   0
            MousePointer    =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PasswordChar    =   ""
            RightToLeft     =   0   'False
            Text            =   ""
            TextType        =   0
            SimulateGrid    =   0   'False
            HighlightText   =   -1  'True
            AllowSpaces     =   -1  'True
            AllUpperCase    =   -1  'True
            HideSelection   =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoTab         =   -1  'True
            AllowPeriods    =   0   'False
         End
         Begin VB.Label lblDolly
            Caption         =   "Dolly:"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   24
            Tag             =   "221"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   2880
         Width           =   2040
         Begin HFCSTextBoxControl.HFCSTextBox txtDolly
            Height          =   285
            Index           =   1
            Left            =   825
            TabIndex        =   9
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483643
            Enabled         =   -1  'True
            ForeColor       =   -2147483640
            BorderStyle     =   1
            DataField       =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataMember      =   ""
            Locked          =   0   'False
            MaxLength       =   0
            MousePointer    =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PasswordChar    =   ""
            RightToLeft     =   0   'False
            Text            =   ""
            TextType        =   0
            SimulateGrid    =   0   'False
            HighlightText   =   -1  'True
            AllowSpaces     =   -1  'True
            AllUpperCase    =   -1  'True
            HideSelection   =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoTab         =   -1  'True
            AllowPeriods    =   0   'False
         End
         Begin VB.Label lblDolly
            Caption         =   "Dolly:"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Tag             =   "221"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   1200
         Width           =   2040
         Begin HFCSTextBoxControl.HFCSTextBox txtDolly
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   40
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483643
            Enabled         =   -1  'True
            ForeColor       =   -2147483640
            BorderStyle     =   1
            DataField       =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            DataMember      =   ""
            Locked          =   0   'False
            MaxLength       =   0
            MousePointer    =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PasswordChar    =   ""
            RightToLeft     =   0   'False
            Text            =   ""
            TextType        =   0
            SimulateGrid    =   0   'False
            HighlightText   =   -1  'True
            AllowSpaces     =   -1  'True
            AllUpperCase    =   -1  'True
            HideSelection   =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoTab         =   -1  'True
            AllowPeriods    =   0   'False
         End
         Begin VB.Label lblDolly
            Caption         =   "Dolly:"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   20
            Tag             =   "221"
            Top             =   240
            Width           =   495
         End
      End
      Begin TrueOleDBGrid70.TDBGrid grdBayDetails
         Height          =   955
         Index           =   1
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1920
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   1693
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   3
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         AllowArrows     =   0   'False
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
        ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid grdBayDetails
         Height          =   955
         Index           =   2
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3720
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   1693
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   3
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         AllowArrows     =   0   'False
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
   End
   Begin HFCSWkEndingDateCntl.WkEndingDtControl ctlWkendingDate
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Weeks           =   0
      Locked          =   0   'False
      BackColor       =   -2147483633
   End
   Begin HFCSRecordCountControl.RecordCount ctlRecordCount
      Height          =   315
      Left            =   12960
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   10040
      Width           =   2235
      _ExtentX        =   3916
      _ExtentY        =   635
   End
   Begin MSComctlLib.ListView ListView1
      Height          =   30
      Left            =   1860
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   10440
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   53
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin HFCSWarningPopup.WarningPopup ctlWarningPopup
      Height          =   735
      Left            =   7440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
   End
   Begin MSComctlLib.Toolbar tbrStd
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   1482
      ButtonWidth     =   979
      ButtonHeight    =   1323
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin HFCSTBManager.HFCSToolbarManager ctlToolbarManager
         Left            =   1200
         Top             =   0
        _ExtentX        =   2646
         _ExtentY        =   1217
      End
      Begin VB.Frame Frame2
         Caption         =   "Frame1"
         BeginProperty Font
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   6375
      End
   End
   Begin VB.Frame fraJobInformation
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   13935
      Begin VB.ComboBox cboDow
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin TrueOleDBGrid70.TDBGrid grdMultLegView
         Height          =   2295
         Left            =   10560
         TabIndex        =   4
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4048
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblDriverNotified
         Caption         =   "Driver Notified:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   11160
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblEloc
         Caption         =   "Eloc/Cn:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   12000
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTractorNumber
         Caption         =   "Tractor Number:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   9120
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblDriverName
         Caption         =   "Driver Name:"
         Height          =   255
         Left            =   6480
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblJobNumber
         Caption         =   " Job Number:"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblWkEnding
         Caption         =   "WE Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblDow
         Caption         =   "DOW:"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Image imgLock
      Height          =   225
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDeleteRecord
      Caption         =   "  Delete Load (Ctrl + Del) "
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   8280
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Menu mnuFile
      Caption         =   "&File"
      Begin VB.Menu mnuPrint
         Caption         =   "&Print"
         Begin VB.Menu mnuPrint_To_Printer
            Caption         =   "Printer"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuPrint_To_Spreadsheet
            Caption         =   "Spreadsheet"
         End
         Begin VB.Menu mnuPrint_To_File
            Caption         =   "File"
         End
      End
      Begin VB.Menu mnuReset
         Caption         =   "Reset Form "
      End
      Begin VB.Menu mnuClose
         Caption         =   "Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuFunctionKeys
      Caption         =   "Function &Keys"
      Begin VB.Menu mnuF2TrailerMessage
         Caption         =   "Tlr / Ld Message"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuF3JobSchedules
         Caption         =   "Driver Schedule"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSchByLds
         Caption         =   "Schedule By Load"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuF4MultiBayDisplay
         Caption         =   "Multi-Bay"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuF5QuickSort
         Caption         =   "Quick Search"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuF6AddUpdateLeg
         Caption         =   "Add/Update Leg"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuF7MultipleLoad
         Caption         =   "Multi-Load"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuPDLdsDisplay
         Caption         =   "Predispatch Loads Display"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPDLdsDisplay
         Caption         =   "Non-Predispatch Lods Display"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuDolliesDisplay
         Caption         =   "Dollies On Property"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuF11CancelPredispatch
         Caption         =   "Cancel Predispatch"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuF10Accept
         Caption         =   "Accept"
      End
   End
   Begin VB.Menu mnuView
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar
         Caption         =   "&Toolbar"
         Begin VB.Menu mnuStandardToolbar
            Caption         =   "Standard &Toolbar"
         End
      End
   End
   Begin VB.Menu mnuHelp
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmPDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '<Revision History>
    '<WR1024><ADID: dmx7plm><Date: 3 March 2010>
    '<Need to enable F4 and F5 bottons on load of screen 13>
Option Explicit
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private mCtlbayfirst As Boolean
Private mbtab   As Boolean
Private iCounter As Integer
Private gpData As PdStruct
Private gbBayReached As Boolean
Private mbElocValid As Boolean
Private mbClearData As Boolean
Private sPreJobNr   As String
Private rsRecordset As ADODB.Recordset
Private sLastBayNr(1 To 3) As String
 
Private sPrevJobNr As String
Private txtLastDriverName As String
Private txtLastEloc As String
Private txtLastTractorNumber As String
Private m_bEditTrailerMessage As Boolean
Private m_bEditLoadMessage As Boolean
Private m_sDllyEntityKey(2) As String
Private WithEvents oNotify As HFCSWSInformation.clsNotification
Attribute oNotify.VB_VarHelpID = -1
Private bBalNrValidated(2)  As Boolean
Private bBayCompeleted(2)  As Boolean
Private bDollyValidated(2) As Boolean
Private bctlBayValidateComplete(2) As Boolean
Public bIntxtDriverNa As Boolean
Private m_bInitializing As Boolean
Private m_bDropDownTlrTyp(2) As Boolean
Private bExitScreen As Boolean
Private bESC        As Boolean
Private bSaveData   As Boolean
Private m_sChildName  As String
Const MAX_SLIC_LENGTH As Integer = 8
Const NO_ELOC   As String = "No Eloc"
' TT000394
Const LIST_NO_GRID As String = "List_No_Grids"
Const LIST_GRIDS As String = "LIST_Grids"
Const MAX_RECOVER_QUERIES           As Long = 3
 
Private oMycomm As HFCSPredispatchDriver.clsComm
Private mydolly As String
Dim WithEvents oshell As HFCSWSInformation.clsNotification
Attribute oshell.VB_VarHelpID = -1
Private m_HFCSDB As HFCSDatabaseConnection.DbConn
Private ctlControls As Control
Private xaMultiLegArray As XArrayDB
Private m_rsLegLoadInfo As ADODB.Recordset
Private rsMultiLeg As ADODB.Recordset
Private m_bClickCancelPD As Boolean
Private m_bGridGotFocus As Boolean
Private JobSysNrOid As String
'TT#3330 - app2dmw - 3/25/20013
Private jobLegsUsedUp As Boolean
Private sRouteCodeMessage(3) As String
Private sRouteCodeChanged(3) As String
Private bTrailerError As Boolean
Private bLoadRoutingError As Boolean
Private bEscPressedOnError As Boolean
Private bSetFocusonBay As Boolean 'used when cancelling load routings
Private bGridGotFocus As Boolean
Private bRoutingCanceled As Boolean
Private lRecoverQueries As Long
 
Private Sub SetupGrid()
    Dim oUtils As HFCSGlobalUtils.clsGridFunctions
    Dim i As Integer
    On Error GoTo Error_Handler
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetupGrid()- Begin")
    End If
   
    Set oUtils = New HFCSGlobalUtils.clsGridFunctions
   
    For i = 0 To 2
        grdBayDetails(i).HeadLines = 2
        Call oUtils.GridInitialize(grdBayDetails(i), ReadOnly)
      '  Call oUtils.Set_Grid_Properties(grdBayDetails(i), ReadOnly)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_TRAILER, "Trailer", dbgNone, dbgNumberOfColumns, 2, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_LOAD, "Load Name", dbgNone, dbgNumberOfColumns, 5, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_LDCD, "  ", dbgNone, dbgNumberOfColumns, 2, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_VOLUME, "Volume", dbgNone, dbgNumberOfColumns, 2, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_TAG, "  ", dbgNone, dbgNumberOfColumns, 1, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_DATE, "Load Dates", dbgNone, dbgNumberOfColumns, 2, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_DEP_TIME, "  ", dbgNone, dbgNumberOfColumns, 1, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_REMARKS, "Remarks", dbgNone, dbgNumberOfColumns, 1, dbgCenter)
        Call oUtils.GridAddSplit(grdBayDetails(i), GRD1_SPL_BAY, "  ", dbgNone, dbgNumberOfColumns, 3, dbgCenter)
   
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_BAY, "Bay #", GRD1_SPL_BAY, 569.7638, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_POS, "Pos", GRD1_SPL_BAY, 345.2599, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_PD, "PD", GRD1_SPL_BAY, 345.2599, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_TRAILER, "#", GRD1_SPL_TRAILER, 1574.929, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_TYPE, "Type", GRD1_SPL_TRAILER, 555.02036, "", True, False, False, False, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_ORIG, "Orig", GRD1_SPL_LOAD, 1005.165, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_OS, "S", GRD1_SPL_LOAD, 285.1654, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_DEST, "Dest", GRD1_SPL_LOAD, 1005.165, "", True, False, False, False, dbgCenter, dbgCenter, Editable, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_DS, "S", GRD1_SPL_LOAD, 285.1654, "", True, False, False, False, dbgCenter, dbgCenter, Editable, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_SEQ, "Sq", GRD1_SPL_LOAD, 315.2126, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_LDCD, "Ld Cd", GRD1_SPL_LDCD, 390.0473, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_RI, "RI", GRD1_SPL_LDCD, 360, "", True, False, False, True, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_PCS, "Pcs", GRD1_SPL_VOLUME, 750.0473, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_PER, "%", GRD1_SPL_VOLUME, 374.7402, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_TAG, "Tag", GRD1_SPL_TAG, 360, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_CREATE, "Create", GRD1_SPL_DATE, 750.0473, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_DUE, "Due", GRD1_SPL_DATE, 750.0473, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_DEP_TM, "Sch Dep Time", GRD1_SPL_DEP_TIME, 734.7402, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
       
        Call oUtils.GridAddColumn(grdBayDetails(i), GRD1_COL_REMARKS, "Due", GRD1_SPL_REMARKS, 3539.906, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, False)
   
        grdBayDetails(i).AllowRowSelect = False
        grdBayDetails(i).ReBind
    Next i
   
       
    Set oUtils = Nothing
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetupGrid()- End")
    End If
    Exit Sub
Error_Handler:
glErrNum = 400
gsError = "SetupGrid"
'the object has disconnect from its clients grid error, retry the grid rebind
If Err.Number = -2147417848 Then
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD SetupGrid()- received error -2147417848")
    End If
   
    If RecoverGridError(grdBayDetails(i)) = True Then
        Resume Next
    End If
End If
Call oProc.update_error_object(Me, gsError)
   
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                           IIf(Err.Number <> 0, Err.Number, glErrNum), _
                           Err.Description & " Module:" & gsError, _
                           oProc, _
                           ERROR_MSG, _
                           FEEDER_DISPATCH_DRIVER) Then
Set oEventlog = Nothing
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'Unload Me
End If
End Sub
 
Private Sub cboDow_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD cboDow_KeyPress - Begin: Key Press: " & KeyAscii)
    End If
    For i = 0 To cboDow.ListCount - 1
       If UCase$(Chr$(KeyAscii)) = Left$(cboDow.List(i), 1) Then
            If cboDow.Text <> cboDow.List(i) Then
                cboDow.Text = cboDow.List(i)
                Exit For
            End If
        End If
    Next i
 
    If (KeyAscii = oKeyDefs.Enter_Key) Then
       EnterAsTab KeyAscii
       KeyAscii = KeyToUpperCase(KeyAscii)
       If g_iDebug = 13 Then
       Call InfoLog("frmPDD cboDow_KeyPress - End KeyAscii If statement")
    End If
       Exit Sub
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD cboDow_KeyPress - End: Key Press: " & KeyAscii)
    End If
End Sub
'
'Private Sub cboDow_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim i As Integer
'    If g_iDebug = 13 Then
'       Call InfoLog("frmPDD cboDow_KeyUp - Begin")
'    End If
'
'If KeyCode = 38 Then
'   For i = 0 To cboDow.ListCount - 1
'      If cboDow.Text = cboDow.List(i) Then
'            If i > 0 Then
'               cboDow.Text = cboDow.List(i - 1)
'            End If
'            If g_iDebug = 13 Then
'                Call InfoLog("frmPDD cboDow_KeyUp - End - KeyCode = 38 If Statement")
'            End If
'            Exit Sub
'      End If
'   Next i
'ElseIf KeyCode = 40 Then
'   For i = 0 To cboDow.ListCount - 1
'      If cboDow.Text = cboDow.List(i) Then
'            If i < cboDow.ListCount - 1 Then
'               cboDow.Text = cboDow.List(i + 1)
'            End If
'            Exit For
'      End If
'   Next i
'End If
'    If g_iDebug = 13 Then
'        Call InfoLog("frmPDD cboDow_KeyUp - End")
'    End If
'End Sub
 
Private Sub chkDriverNotified_Click()
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkDriverNotified_Click() - Begin")
    End If
    pInst.PDInfo.szFdrDvrInfNtfIr = IIf(chkDriverNotified.Value = vbChecked, IR_TRUE, IR_FALSE)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkDriverNotified_Click() - End")
    End If
End Sub
 
Private Sub chkExtraDolly_GotFocus(Index As Integer)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkExtraDolly_GotFocus - Begin - Index: " & Index)
    End If
   
    If (Index <> ictlIndex) Then
         ictlIndex = Index
    End If
   
    chkExtraDolly(Index).BackColor = vbGreen
   
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkExtraDolly_GotFocus - End - Index: " & Index)
    End If
End Sub
 
Private Sub chkExtraDolly_LostFocus(Index As Integer)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkExtraDolly_LostFocus - Begin - Index: " & Index)
    End If
   
    chkExtraDolly(Index).BackColor = frmPDD.BackColor
   
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkExtraDolly_LostFocus - End - Index: " & Index)
    End If
End Sub
 
Private Sub chkExtraDolly_Validate(Index As Integer, Cancel As Boolean)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkExtraDolly_Validate - Begin - Index: " & Index)
    End If
   
    'TT00394
    If Not IsUserShiftTabbing Then
        ctlBayNr.Item(Index + 1).Visible = True
        ctlBayNr(Index + 1).Enabled = True
        ctlBayNr.Item(Index + 1).SetFocus
    End If
   
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD chkExtraDolly_Validate - End - Index: " & Index)
    End If
End Sub
 
Private Sub ctlBayNr_GotFocus(Index As Integer)
On Error GoTo ErrorExit
'<WR01024><ADID: dmx7plm><Enable F4 and F5 buttons on load of Screen 13 from Feeder Shell> - Start
Dim lHwnd As Long
If g_iDebug = 13 Then
    Call InfoLog("frmPDD ctlBayNr_GotFocus - Begin - Index: " & Index)
End If
 
'If m_bGridGotFocus = True Then grdBayDetails(iBayIndex).SetFocus
ctlBayNr(Index).BackgroundColor = vbGreen
ictlIndex = Index
 
 
If OGlobalFormNames Is Nothing Then Set OGlobalFormNames = New GLOBALDEFS.clsFormNames
 
If bClickedF4 = True Then lHwnd = GetWindowHandle(OGlobalFormNames.StartBaySearch)
If bClickedF5 = True Then lHwnd = GetWindowHandle(OGlobalFormNames.QuickSearchF5)
 
If (lHwnd <> 0) Then Call SendMessage(lHwnd, WM_MCL_CHANGEWRITE, 0, 0)
'<WR01024><ADID: dmx7plm><Enable F4 and F5 buttons on load of Screen 13 from Feeder Shell> - End
 
Dim bEnF2 As Boolean, bEnF7 As Boolean
 
Dim i As Integer
    'Only enable the one that following this bay
   ' If bSetFocusonBay = False Then
    txtDolly.Item(Index).Enabled = True
   ' Else
   '     grdBayDetails(Index).TabStop = False
   '     bSetFocusonBay = False
   ' End If
   
    TextEnd.TabStop = False
    TextEnd.Enabled = False
   
    bExitScreen = False
    bIntxtDriverNa = False 'Now it is OK to open 890 and 891
   
    'Remove all the loads that are not on property, but still show on screen.
    For i = 1 To pInst.PDInfo.hPDHll.Count
         If pInst.PDInfo.hPDHll.Count > 0 And i <= pInst.PDInfo.hPDHll.Count Then
         Else
            ctlBayNr.Item(0).Visible = True
            ctlBayNr.Item(0).Enabled = True
            ctlBayNr.Item(0).SetFocus
         End If
    Next i
      
    If iMyRc = MBID_CANCEL Then
       ctlBayNr.Item(0).Enabled = False
       txtJobNr.Enabled = True
       txtJobNr.SetFocus
    ElseIf iMyRc = UMBID_DELETE Then
       ctlBayNr.Item(0).Visible = True
       ctlBayNr.Item(0).Enabled = True
       ctlBayNr.Item(0).SetFocus
       iMyRc = 0
    ElseIf iMyRc = DB_EOF Then
       ctlBayNr.Item(0).Visible = True
       ctlBayNr.Item(0).Enabled = True
       ctlBayNr.Item(0).SetFocus
       iMyRc = 0
    End If
 
    If Not bReLoad Then
       pInst.icdl = Index + 1
    End If
    iBayIndex = Index
 
    chkDriverNotified.Enabled = False
 
    If Len(grdBayDetails.Item(Index).Columns(GRD1_COL_TRAILER).Text) = 0 And _
       Len(grdBayDetails.Item(Index).Columns(GRD1_COL_DEST).Text) = 0 Then
       bLostFocus = False
       lblDeleteRecord.Visible = False
    End If
 
    If Not bBalNrValidated(Index) Then
       If Index = 0 Then
           SetAvailMenuAndTool MAT_BASE_DATA0, False, False
       Else
           SetAvailMenuAndTool mat_bay, False, False
       End If
    End If
    If Len(ctlBayNr.Item(Index).Text) > 0 And pInst.PDInfo.hPDHll.Count > 0 And pInst.PDInfo.hPDHll.Count > Index Then '5/24
'        If pInst.PDInfo.hPDHll.Item(Index + 1).szMultLdIr = IR_TRUE And (pInst.PDInfo.hPDHll.Item(Index + 1).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(Index + 1).lLdMsgId > 0) Then
'           SetAvailMenuAndTool mat_bay, True, True
'        ElseIf pInst.PDInfo.hPDHll.Item(Index + 1).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(Index + 1).lLdMsgId > 0 Then
'           SetAvailMenuAndTool mat_bay, True, False
        If pInst.PDInfo.hPDHll.Item(Index + 1).szMultLdIr = IR_TRUE Then
           SetAvailMenuAndTool mat_bay, True, True
        Else
           SetAvailMenuAndTool mat_bay, True, False
        End If
    Else
        If Index = 0 Then
           If iMyRc <> MBID_CANCEL Then   'Added for TSAT 2125
             SetAvailMenuAndTool mat_bay, False, False
           End If
        Else
           SetAvailMenuAndTool mat_bay, False, False
        End If
    End If
    pInst.icdl = pInst.PDInfo.hPDHll.Count
    ctlBayNr.Item(Index).ZOrder 0
    If pInst.PDInfo.hPDHll.Count > 0 And pInst.PDInfo.hPDHll.Count >= pInst.icdl Then
        If Len(pInst.PDInfo.hPDHll(pInst.icdl).szBayNr) > 0 Then
           bEnF2 = True
        Else
           bEnF2 = (pInst.PDInfo.hPDHll(pInst.icdl).lLdMsgId <> 0 Or _
                                                pInst.PDInfo.hPDHll(pInst.icdl).lTlrMsgId <> 0)
        End If
        bEnF7 = (pInst.PDInfo.hPDHll(pInst.icdl).szMultLdIr = IR_TRUE)
    Else
        bEnF2 = False
        bEnF7 = False
    End If
 
    If pInst.PDInfo.hPDHll.Count > Index Then
        If Not IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) And Not IsNull(pInst.PDInfo.SegmentSystemNumberOID) Then
           If Trim$(pInst.PDInfo.SegmentSystemNumberOID) = Trim$(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Then
               ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, True
               mnuF11CancelPredispatch.Enabled = True
           Else
               ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
               mnuF11CancelPredispatch.Enabled = False
           End If
        Else
           ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
           mnuF11CancelPredispatch.Enabled = False
        End If
    Else
        ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
        mnuF11CancelPredispatch.Enabled = False
    End If
   
    If Not mbElocValid Then
        If ClearData Then
           ClearPddData
           txtJobNr.Enabled = True
           txtJobNr.SetFocus
        End If
       
'Removed in order to correct Defect 501; this logic stated that if the user came from an arrival screen once they got focus
'in the ctlBayNr to set the focus back to the grdMultLegView. This is not the correct flow for the form and caused runtime errors to occur - JLW8PGC
'        If pInst.bArriveFlag Then
'           grdMultLegView.SetFocus ' txtEloc.SetFocus
'        Else
'        End If
    Else
                If g_iDebug = 13 Then
                        Call InfoLog("frmPDD ctlBayNr_GotFocus - Eloc valid - End - Index: " & Index)
                End If
                Exit Sub
                End If
 
                If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_GotFocus - End - Index: " & Index)
    End If
   
    'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
'    bRoutingCanceled = False
Exit Sub
 
ErrorExit:
glErrNum = 400
gsError = "ctlBayNr_GotFocus"
Call oProc.update_error_object(Me, gsError)
 
    Screen.MousePointer = vbNormal
        If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                 oProc, _
                                 ERROR_MSG, _
                                 FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Sub
 
Private Sub ctlBayNr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_KeyDown - Begin - Index: " & Index)
    End If
    bDuplicateLoad = False
    If bBayCompeleted(Index) Then
       KeyCode = 0
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_KeyDown - End - Index: " & Index)
    End If
End Sub
 
Private Sub ctlBayNr_KeyPress(Index As Integer, KeyAscii As Integer)
   
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_KeyPress - Begin - Index: " & Index)
    End If
    If bBayCompeleted(Index) Then
       KeyAscii = 0
       If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_KeyPress - End - If bBayCompeleted(Index) - Index: " & Index)
       End If
       Exit Sub
    End If
   
  If KeyAscii >= 48 And KeyAscii <= 57 Then
     bDisableF10 = True
     mnuF10Accept.Enabled = False
     ctlToolbarManager.buttonEnabled BT_F10_ACCEPT, mnuF10Accept.Enabled
 
     'When arrival load is deleted, then delete the dolly which came with the trailer.
     If bFromGTForm(Index) And Len(txtDolly.Item(Index).Text) > 0 Then
        txtDolly.Item(Index).Text = vbNullString
     End If
     bFromF4F5 = False
     bDatafromMatchingScreen(Index) = False
     bDataFromEloc(Index) = False
     bFromGTForm(Index) = False
    
     If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
     End If
  End If
 
  If pInst.bArriveFlag And _
     bBayCompeleted(iBayIndex) And Me.ActiveControl Is ctlBayNr.Item(iBayIndex) Then
       KeyAscii = 0
       If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_KeyPress - End - If pInst.bArriveFlag And... - Index: " & Index)
       End If
       Exit Sub
  End If
 
  If (KeyAscii = oKeyDefs.Enter_Key) Then
     EnterAsTab KeyAscii
     KeyAscii = KeyToUpperCase(KeyAscii)
     If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_KeyPress - End - If (KeyAscii = oKeyDefs.Enter_Key) - Index: " & Index)
       End If
     Exit Sub
  End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_KeyPress - End - Index: " & Index)
    End If
       
End Sub
 
Private Sub ctlBayNr_LostFocus(Index As Integer)
1    On Error GoTo ErrorHandler
Dim lErrline As Integer
Dim bCancel As Boolean
 
'<WR01024><ADID: dmx7plm><Enable F4 and F5 buttons on load of Screen 13 from Feeder Shell> - Start
2 Dim lHwnd As Long
3    If g_iDebug = 13 Then
4        Call InfoLog("frmPDD ctlBayNr_LostFocus - Begin - Index: " & Index)
5    End If
  bCancel = False
    If Me.Enabled = False Then Exit Sub
6 ctlBayNr(Index).BackgroundColor = vbWhite
7 ictlIndex = Index
'clear out member variable for F4 and F5 buttons
8 If m_sChildName = OGlobalFormNames.QuickSearchF5 Or m_sChildName = OGlobalFormNames.MultiBaySearchF4 Then
9   m_sChildName = vbNullString
10 End If
11 If m_bClickCancelPD = True Then
12    If g_iDebug = 13 Then
13            Call InfoLog("frmPDD ctlBayNr_LostFocus - End - If m_bClickCancelPD = True - Index: " & Index)
14    End If
15    Exit Sub
16 End If
 
17 If OGlobalFormNames Is Nothing Then Set OGlobalFormNames = New GLOBALDEFS.clsFormNames
 
18 If bClickedF4 = True Then lHwnd = GetWindowHandle(OGlobalFormNames.StartBaySearch)
19 If bClickedF5 = True Then lHwnd = GetWindowHandle(OGlobalFormNames.QuickSearchF5)
 
20    If g_iDebug = 13 Then
21            Call InfoLog("frmPDD ctlBayNr_LostFocus - before SendMessage - If m_bClickCancelPD = True - Index: " & Index)
22    End If
23 If (lHwnd <> 0) Then Call SendMessage(lHwnd, WM_MCL_CHANGEWRITE, 0, 0)
'<WR01024><ADID: dmx7plm><Enable F4 and F5 buttons on load of Screen 13 from Feeder Shell> - End
24    If g_iDebug = 13 Then
25            Call InfoLog("frmPDD ctlBayNr_LostFocus - after SendMessage - If m_bClickCancelPD = True - Index: " & Index)
26    End If
     'Only when the previous Bay lost focus, then enable the next bay.
27   If (Index < 2) Then
28      ctlBayNr.Item(Index + 1).Enabled = True
29      ctlBayNr.Item(Index + 1).TabStop = True
30   End If
 
31   If bESC Then
32       bESC = False
       'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
33       If Not pInst Is Nothing Then
34           If Not pInst.bArriveFlag And Not bFromGtF8 Then
35              txtJobNr.Enabled = True
36              txtJobNr.SetFocus
37           ElseIf Me.Enabled = True Then
38              grdMultLegView.Enabled = True
39              txtDriverName.Enabled = True
40              txtDriverName.SetFocus
    '          txtEloc.Enabled = True
        '      txtTractorNumber.Enabled = True
        '      txtTractorNumber.SetFocus
41           End If
42       End If
43       If g_iDebug = 13 Then
44           Call InfoLog("frmPDD ctlBayNr_LostFocus -End - If bESC = true - Index: " & Index)
45       End If
46       Exit Sub
47   End If
 
48   If bFromGTForm(Index) And Len(grdBayDetails.Item(Index).Columns(GRD1_COL_TRAILER).Text) > 0 Then
49      If bLostFocus Then
50         If Not (IsUserShiftTabbing Or m_bClickCancelPD) Then
500          If bctlBayValidateComplete(Index) = False And bPredispatchEmpty(Index) = False Then
501       '      Call ctlBayNr_Validate(Index, bCancel)
                Me.Enabled = True
                ctlBayNr(Index).Visible = True
                ctlBayNr(Index).Enabled = True
                If ctlBayNr(Index).Enabled = True Then ctlBayNr(Index).SetFocus
                Exit Sub
502          End If
51           bBayCompeleted(Index) = True
      '     txtDolly.Item(Index).SetFocus
52           If Index = 0 Then
53              txtDolly.Item(Index).Enabled = True
54              txtDolly.Item(Index).TabStop = True
55              chkExtraDolly(Index).Enabled = True
56              chkExtraDolly(Index).TabStop = True
57              chkExtraDolly(1).Enabled = False
58              chkExtraDolly(1).TabStop = False
59           ElseIf Index = 1 Then
60              If Len(Trim$(txtDolly(1).Text)) = 0 Then
61                chkExtraDolly(Index).Value = chkExtraDolly(0).Value
62              End If
             
63              If Not (Len(Trim$(txtDolly(0).Text)) > 0 And Len(Trim$(txtDolly(1).Text)) = 0) Then
64                chkExtraDolly(0).Value = 0
65                chkExtraDolly(0).Enabled = False
66                chkExtraDolly(0).TabStop = False
67              End If
 
68              chkExtraDolly(Index).Enabled = True
69              chkExtraDolly(Index).TabStop = True
           
70              If (Len(Trim$(txtDolly(1).Text)) > 0 Or chkExtraDolly(0).Value = 1) Then
71                chkExtraDolly(Index).Enabled = False
72                chkExtraDolly(Index).TabStop = False
73              End If
74           End If
75           If g_iDebug = 13 Then
76              Call InfoLog("frmPDD ctlBayNr_LostFocus - End - Not User Shift Tabbing - Index: " & Index)
77           End If
78           Exit Sub
79         ElseIf IsUserShiftTabbing And m_bClickCancelPD = False Then
503          If bctlBayValidateComplete(Index) = False And bPredispatchEmpty(Index) = False Then
504       '      Call ctlBayNr_Validate(Index, bCancel)
                Me.Enabled = True
                ctlBayNr(Index).Visible = True
                ctlBayNr(Index).Enabled = True
                If ctlBayNr(Index).Enabled = True Then ctlBayNr(Index).SetFocus
                Exit Sub
505          End If
80           If Index > 0 Then
81              txtDolly.Item(Index - 1).Enabled = True
82              txtDolly.Item(Index - 1).SetFocus  'goto previous dolly 2/2/06
83           Else
84              ctlBayNr.Item(Index).Visible = True
85              ctlBayNr.Item(Index).Enabled = True
86              ctlBayNr.Item(Index).SetFocus  'Stay 2/2/06
87           End If
88           bBayCompeleted(Index) = True
89           If Index = 0 Then
90              chkExtraDolly(Index).Enabled = True
91              chkExtraDolly(Index).TabStop = True
92              chkExtraDolly(1).Enabled = False
93              chkExtraDolly(1).TabStop = False
94           ElseIf Index = 1 Then
95              If Len(Trim$(txtDolly(1).Text)) = 0 Then
96                chkExtraDolly(Index).Value = chkExtraDolly(0).Value
97              End If
             
98              If Not (Len(Trim$(txtDolly(0).Text)) > 0 And Len(Trim$(txtDolly(1).Text)) = 0) Then
99                chkExtraDolly(0).Value = 0
100                chkExtraDolly(0).Enabled = False
101                chkExtraDolly(0).TabStop = False
102              End If
103               chkExtraDolly(Index).Enabled = True
104              chkExtraDolly(Index).TabStop = True
105
106              If (Len(Trim$(txtDolly(1).Text)) > 0 Or chkExtraDolly(0).Value = 1) Then
107                chkExtraDolly(Index).Enabled = False
108                chkExtraDolly(Index).TabStop = False
109              End If
110           End If
           
111           If g_iDebug = 13 Then
112              Call InfoLog("frmPDD ctlBayNr_LostFocus -End - User Shift Tabbing - Index: " & Index)
113           End If
114           Exit Sub
115         End If
116      End If
117   End If
 
118   If bMultiForm Then
119      If frmSelMultiLds.Visible Then
120         If frmSelMultiLds.grdMultiLoads.Enabled = False Then frmSelMultiLds.Enabled = True
121         frmSelMultiLds.grdMultiLoads.SetFocus
122         If g_iDebug = 13 Then
123            Call InfoLog("frmPDD ctlBayNr_LostFocus - End - bMultiForm visible - Index: " & Index)
124         End If
125         Exit Sub
126      Else
127         If g_iDebug = 13 Then
128            Call InfoLog("frmPDD ctlBayNr_LostFocus - End - bMultiForm not visible - Index: " & Index)
129         End If
130         Exit Sub
131      End If
132   ElseIf bAtGTForm Then
133      If frmLdsAtGate.Visible Then
134         If frmLdsAtGate.grdLoads.Enabled = False Then frmLdsAtGate.grdLoads.Enabled = True
135         frmLdsAtGate.grdLoads.SetFocus
136         If g_iDebug = 13 Then
137            Call InfoLog("frmPDD ctlBayNr_LostFocus - End - bAtGtForm visible - Index: " & Index)
138         End If
139         Exit Sub
140      Else
141         If g_iDebug = 13 Then
142            Call InfoLog("frmPDD ctlBayNr_LostFocus - End - bAtGtForm not visible - Index: " & Index)
143         End If
144         Exit Sub
145      End If
146   ElseIf bMatchingForm Then
147      If frmMatchingLds.Visible Then
148         If frmMatchingLds.grdLoads.Enabled = False Then frmMatchingLds.grdLoads.Enabled = True
149         frmMatchingLds.grdLoads.SetFocus
150         If g_iDebug = 13 Then
151            Call InfoLog("frmPDD ctlBayNr_LostFocus - End - bMatchingForm visible - Index: " & Index)
152         End If
153         Exit Sub
154      Else
155         If g_iDebug = 13 Then
156            Call InfoLog("frmPDD ctlBayNr_LostFocus -End - bMatchingForm not visible - Index: " & Index)
157         End If
158         Exit Sub
159      End If
160   End If
 
161   If (Len(ctlBayNr.Item(Index).Text) = 0 And grdBayDetails(Index).Columns(GRD1_COL_PD).Value = HFCS_INDICATOR_FALSE) Or (bBalNrValidated(Index) = False) Then
162       If iMyRc <> MBID_CANCEL Then
163          If (Trim(ctlBayNr.Item(Index).Text) <> "" And Not (Me.ActiveControl Is grdBayDetails(Index))) Then
164             ctlBayNr.Item(Index).Visible = True
165             ctlBayNr.Item(Index).Enabled = True
                If ctlBayNr.Item(Index).Enabled = True Then
                    ctlBayNr.Item(Index).SetFocus
                Else
                    Exit Sub
                End If
167          End If
         
168          If Trim(ctlBayNr.Item(Index).Text) = "" And bPredispatchEmpty(Index) = False And bPredispatchEmptyWithTrailer(Index) = False And Not IsUserShiftTabbing Then
169             ctlBayNr.Item(Index).Visible = True
170             ctlBayNr.Item(Index).Enabled = True
171             ctlBayNr.Item(Index).SetFocus
172          ElseIf Not IsUserShiftTabbing Then
506             If bctlBayValidateComplete(Index) = False And bPredispatchEmpty(Index) = False Then
507        '         Call ctlBayNr_Validate(Index, bCancel)
                     Me.Enabled = True
                     ctlBayNr(Index).Visible = True
                     ctlBayNr(Index).Enabled = True
                     If ctlBayNr(Index).Enabled = True Then ctlBayNr(Index).SetFocus
                     Exit Sub
508             End If
173             grdBayDetails(Index).AllowUpdate = True
174             grdBayDetails(Index).Enabled = True
175             grdBayDetails(Index).SetFocus
176             ctlBayNr.Item(Index).Enabled = False
177             ctlBayNr.Item(Index).Visible = False
178             bBalNrValidated(Index) = True
179             bBayCompeleted(Index) = True
180             If Index = 0 Then
181                chkExtraDolly(Index).Enabled = True
182                chkExtraDolly(Index).TabStop = True
183                chkExtraDolly(1).Enabled = False
184                chkExtraDolly(1).TabStop = False
185             ElseIf Index = 1 Then
186                If Len(Trim$(txtDolly(1).Text)) = 0 Then
187                  chkExtraDolly(Index).Value = chkExtraDolly(0).Value
188                End If
                   
189                If Not (Len(Trim$(txtDolly(0).Text)) > 0 And Len(Trim$(txtDolly(1).Text)) = 0) Then
190                  chkExtraDolly(0).Value = 0
191                  chkExtraDolly(0).Enabled = False
192                  chkExtraDolly(0).TabStop = False
193                End If
194                chkExtraDolly(Index).Enabled = True
195                chkExtraDolly(Index).TabStop = True
                 '  End If
196                If (Len(Trim$(txtDolly(1).Text)) > 0 Or chkExtraDolly(0).Value = 1) Then
197                  chkExtraDolly(Index).Enabled = False
198                  chkExtraDolly(Index).TabStop = False
199                End If
200             End If
201             pInst.bPreDispatched(Index + 1) = True
202            End If
              
        '  If Index = 0 Then  'Bay# is blank, shift+tab and tab, focus will not move.
        '     ctlBayNr.Item(Index).SetFocus
        '     mCtlbayfirst = True
        '  End If
203       End If
204   Else
205       If bLostFocus Then
206          If Index = 0 Then  'Bay# is not blank, if focus is on the first Bay, shift+tab, will stay in where is it now.
207              If Not (IsUserShiftTabbing Or m_bClickCancelPD) Then
              '    SetAvailMenuAndTool mat_bay, True, False  '19/12 QS
208                  bLostFocus = False
209              Else
210                  ctlBayNr.Item(Index).Visible = True
211                  ctlBayNr.Item(Index).Enabled = True
212                  ctlBayNr.Item(Index).SetFocus
213              End If
214          ElseIf Index = 2 Then
         '    bSaveData = True
215          Else
            '  SetAvailMenuAndTool mat_bay, True, False  '19/12 QS
216              bLostFocus = False
217          End If
          'added for editing Load Routings
218          If Not IsUserShiftTabbing Then
219             If (pInst.PDInfo.hPDHll(Index + 1).SameLoadName = False Or pInst.PDInfo.hPDHll(Index + 1).szLdCd = "E") Then
220                grdBayDetails(Index).AllowUpdate = True
221                grdBayDetails(Index).Enabled = True
222                grdBayDetails(Index).SetFocus
            
223                ctlBayNr.Item(Index).Enabled = False
224                ctlBayNr.Item(Index).Visible = False
225             Else
509                If bctlBayValidateComplete(Index) = False And bPredispatchEmpty(Index) = False Then
510            '        Call ctlBayNr_Validate(Index, bCancel)
                     Me.Enabled = True
                     ctlBayNr(Index).Visible = True
                     ctlBayNr(Index).Enabled = True
                     If ctlBayNr(Index).Enabled = True Then ctlBayNr(Index).SetFocus
                     Exit Sub
511                End If
226                bBayCompeleted(Index) = True
227             End If
           
228          End If
         
229          If Index = 0 Then
230             chkExtraDolly(Index).Enabled = True
231             chkExtraDolly(Index).TabStop = True
232             chkExtraDolly(1).Enabled = False
233             chkExtraDolly(1).TabStop = False
234          ElseIf Index = 1 Then
235             If Len(Trim$(txtDolly(1).Text)) = 0 Then
236                chkExtraDolly(Index).Value = chkExtraDolly(0).Value
237             End If
                   
238             If Not (Len(Trim$(txtDolly(0).Text)) > 0 And Len(Trim$(txtDolly(1).Text)) = 0) Then
239                chkExtraDolly(0).Value = 0
240                chkExtraDolly(0).Enabled = False
241                chkExtraDolly(0).TabStop = False
242             End If
243             chkExtraDolly(Index).Enabled = True
244             chkExtraDolly(Index).TabStop = True
 
245             If (Len(Trim$(txtDolly(1).Text)) > 0 Or chkExtraDolly(0).Value = 1) Then
246                chkExtraDolly(Index).Enabled = False
247                chkExtraDolly(Index).TabStop = False
248             End If
249          End If
         
250          bFromF4F5 = False
251       End If
252   End If
253   gbProcessRetData = False
   If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_LostFocus - End - Index: " & Index)
   End If
Exit Sub
 
ErrorHandler:
 
 
lErrline = Erl
glErrNum = 400
gsError = "ctlBayNr_LostFocus"
Call oProc.update_error_object(Me, gsError)
 
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError & " " & lErrline, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Sub
 
Private Sub ctlBayNr_Validate(Index As Integer, Cancel As Boolean)
Dim bResult As Boolean
Dim iRc As Integer
Dim sCurBayNum As String
Dim i As Integer
Dim sDollyOnBay As String
Dim sBayRangName As String
Dim rsRecordset As ADODB.Recordset
Dim iDbSuccess As Integer
Dim bFound As Boolean
Dim oColBay As HFCSYardObject.Bay
Dim bBayStillExist As Boolean
Dim iBayPosition As Integer
Dim rsBays As ADODB.Recordset
Dim bShowEmpty As Boolean
Dim colEquip As Collection
Dim oTlr As HFCSTrailerObject.Trailer
 
On Error GoTo ErrorExit
 
If g_iDebug = 13 Then
    Call InfoLog("frmPDD ctlBayNr_Validate - Begin - Index: " & Index)
End If
If IsUserShiftTabbing And Index = 0 Then
    Cancel = True
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_Validate - End - If IsUserShiftTabbing And Index = 0 - Index: " & Index)
    End If
    Exit Sub
End If
    bEscPressedOnError = False
    bTrailerError = False
If bPredispatchEmpty(Index) = False Then
    If bPreDispatched = True Or gbLoadPreDispatched = True Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_Validate - End - If bPredispatchEmpty - Index: " & Index)
        End If
        bctlBayValidateComplete(Index) = True
        Exit Sub
    End If
   
    If bDataFromEloc(Index) Or bDatafromMatchingScreen(Index) Then
       bLostFocus = True
    End If
   
    If Not pInst.bArriveFlag And ctlBayNr.Item(Index).Text = GATE_BAY_TYPE Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_Validate - End - If Not pInst.bArriveFlag And... - Index: " & Index)
        End If
       bctlBayValidateComplete(Index) = True
       Exit Sub  'if nothing from arrival, GT is not valid.
    ElseIf bBayCompeleted(Index) And bFromGTForm(Index) Then
       bLostFocus = True
       If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_Validate - End - bBayCompleted true  - Index: " & Index)
       End If
        bctlBayValidateComplete(Index) = True
       Exit Sub
    End If
   
    If Not IsUserShiftTabbing Then
        If (ctlWarningPopup.WarningVisible And ctlBayNr.Item(Index).Text <> sPreBayNum(Index)) Or bMultiBay(Index) Then
            ctlWarningPopup.ClearWarning
        ElseIf ctlBayNr.Item(Index).Text <> sPreBayNum(Index) And Not IsBlank(sPreBayNum(Index)) Then
            bBalNrValidated(Index) = False
        ElseIf bBalNrValidated(Index) And bFromGTForm(Index) Then
            bLostFocus = True
            If g_iDebug = 13 Then
                 Call InfoLog("frmPDD ctlBayNr_Validate - End - bBalNrValidated true  - Index: " & Index)
            End If
            bctlBayValidateComplete(Index) = True
            Exit Sub
        End If
    End If
   
    'If it comes from arriving screen, and it is duplicate, then exit
    If bDuplicateLoad Then
       ctlWarningPopup.ShowWarning ctlBayNr(Index).hwnd, LoadResString(sDuplicateload), 2000
       Cancel = True
       ctlBayNr.Item(Index).Visible = True
       ctlBayNr.Item(Index).Enabled = True
       ctlBayNr.Item(Index).SetFocus
       sPreBayNum(Index) = ctlBayNr.Item(Index).Text
       If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_Validate - End - If bDuplicateLoad - Index: " & Index)
       End If
       Exit Sub
    End If
 
    'Stay in focus when Bay# is empty
    If Len(ctlBayNr.Item(Index).Text) = 0 And Not IsUserShiftTabbing Then
      'FE AAD R3: if bay is empty, we assume the user wants to predispatch an empty trailer
     
      If pInst.PDInfo.hPDHll.Count > Index Then
          If pInst.PDInfo.hPDHll(Index + 1).szLdCd <> "E" And pInst.PDInfo.hPDHll(Index + 1).szLdCd <> "C" And IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Then
            ctlBayNr.Item(Index).Visible = True
            ctlBayNr.Item(Index).Enabled = True
            ctlBayNr.Item(Index).SetFocus
            ctlBayNr.Item(Index).SelectText
            Cancel = True
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD ctlBayNr_Validate - End - If pInst.PDInfo.hPDHll.Count > Index - Index: " & Index)
            End If
            Exit Sub
          End If
      End If
     
      'mve tt00512
      If (pInst.PDInfo.hPDHll.Count > Index) Then
         If IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Or IsEmpty(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Then
            bShowEmpty = True
         Else
            bShowEmpty = False
         End If
      Else
         bShowEmpty = True
      End If
                
      If bShowEmpty Then
        iRc = pdMsgBox(LoadResString(sPreDispatchEmptyTrailerQuestion), _
                       vbQuestion + vbYesNo, _
                       LoadResString(sPreDispatchEmptyTrailerHeader))
 
        If iRc <> vbYes Then ' Don't go for the next movement
            ctlBayNr.Item(Index).Visible = True
            ctlBayNr.Item(Index).Enabled = True
            ctlBayNr.Item(Index).SetFocus
            ctlBayNr.Item(Index).SelectText
            Cancel = True
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD ctlBayNr_Validate - End - If bShowEmpty - Index: " & Index)
            End If
            Exit Sub
        End If
     
        bPredispatchEmpty(Index) = True
        SetColumnLocksForEmptyLoad (Index)
        If pInst.PDInfo.hPDHll.Count <= Index Then
           If Index > 0 Then
             If bPredispatchEmpty(Index - 1) Then
                UpdatePredispatchEmptyData (Index - 1)
                grdBayDetails(Index - 1).AllowUpdate = False
             ' TT000394
            grdBayDetails(Index - 1).Enabled = False
               ' grdBayDetails(Index).TabStop = True
             End If
           End If
 
           Call pInst.CreateEmptyPredispatch(Index)
  
           For i = 1 To pInst.PDInfo.hPDHll.Count
             DisplayLoadInfo pInst.PDInfo.hPDHll.Item(i), i - 1
           Next i
       ' ElseIf pInst.PDInfo.hPDHll.Count > 0 Then
       '    pInst.PDInfo.hPDHll.Item(Index + 1).iSequenceNr = 0
        End If
 
        Call SetGridForEdit(Index)
        ' TT000394
        Call ReTab(LIST_GRIDS) ' "LIST_Grids")
        If (Len(Trim$(ctlBayNr(Index).Text)) = 0) Then
            grdBayDetails(Index).Split = 1
            grdBayDetails(Index).Col = GRD1_COL_TYPE
        End If
        bLostFocus = True
      Else
        SetColumnLocksForEmptyLoad (Index)
        Call SetGridForEdit(Index)
        Call ReTab(LIST_GRIDS) ' "LIST_Grids")
        grdBayDetails(Index).Split = 3
        grdBayDetails(Index).Col = GRD1_COL_RI
      End If
    End If
   
    If bFromF4F5 Then   'Data is from F4 or F5 screens.
       bLostFocus = True
    End If
 
    If bBalNrValidated(Index) And sPreBayNum(Index) = ctlBayNr.Item(Index).Text And Not ctlWarningPopup.WarningVisible Then
       bLostFocus = True
       If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_Validate - End - bBalNrValidated and sPreBayNum = Bay true - Index: " & Index)
       End If
       Exit Sub
    End If
 
    If Len(grdBayDetails.Item(Index).Columns(GRD1_COL_TRAILER).Text) = 0 And _
       Len(grdBayDetails.Item(Index).Columns(GRD1_COL_DEST).Text) = 0 Then
       bLostFocus = False
       lblDeleteRecord.Visible = False
    ElseIf Not bDatafromMatchingScreen(Index) Then
       If pInst.PDInfo.lCount = Index + 1 Then
          bDatafromMatchingScreen(Index) = False
       End If
    End If
 
'    glErrNum = 400
'    gsError = LoadResString(sctlBayNrValidate)
'    Call oProc.update_error_object(Me, gsError)
 
    sCurBayNum = ctlBayNr.Item(Index).Text
 
    'Added checking the ORIG is because there are some loads exist but not on property.
    If IsBlank(sCurBayNum) And Not bBalNrValidated(Index) Then
        If Index >= pInst.PDInfo.hPDHll.Count Then
            SetAvailMenuAndTool mat_bay, False, False
            If IsUserShiftTabbing Then
                If g_iDebug = 13 Then
                    Call InfoLog("shift-tabbing back to dolly.")
                End If
                If Index > 0 Then
                    txtDolly(Index - 1).Enabled = True
                    txtDolly(Index - 1).SetFocus
                Else
                    Cancel = True
                    bBalNrValidated(Index) = False
                    bBayCompeleted(Index) = False
                    Exit Sub
                End If
           '     DoEvents
            End If
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD ctlBayNr_Validate - End - If IsBlank(sCurBayNum) And... - Index: " & Index)
            End If
            Exit Sub
        End If
 
        If IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Then
            SetAvailMenuAndTool mat_bay, False, False
            If IsUserShiftTabbing Then
              '  Cancel = False
                If g_iDebug = 13 Then
                    Call InfoLog("shift-tabbing back to dolly.")
                End If
                If Index > 0 Then
                '    txtDolly(Index - 1).SetFocus
                '    DoEvents
                Else
                    Cancel = True
                    bBayCompeleted(Index) = False
                    bBalNrValidated(Index) = False
                    Exit Sub
                End If
            End If
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD ctlBayNr_Validate - End - If IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) - Index: " & Index)
            End If
            Exit Sub
        End If
    End If
 
    If Not ctlBayNr.Item(Index).Text = GATE_BAY_TYPE And Not IsBlank(sCurBayNum) Then
        If Not LockBayForPredispatch(Index) Then
            ctlBayNr.Item(Index).Visible = True
            ctlBayNr.Item(Index).Enabled = True
            ctlBayNr.Item(Index).SetFocus
            ctlBayNr.Item(Index).SelectText
            Cancel = True
 
            If pInst.raLockedBay(Index + 1) <> 0 Then
                ' Unlock the Bay
                pInst.raLockedBay(Index + 1) = 0
            End If
            sLastBayNr(Index + 1) = ""
            If g_iDebug = 13 Then
                 Call InfoLog("frmPDD ctlBayNr_Validate - End - LockBayForPredispatch - Index: " & Index)
            End If
            Exit Sub
        End If
    End If
 
    ' If a bay number has been entered
    ' by the operator, we need to
    ' validate it
    sLastBayNr(Index + 1) = sCurBayNum
 
    'Arrival tractor only, bay # can not be "GD01"
    If ctlBayNr.Item(Index).Text = GATE_BAY_TYPE And bArrivalTractorOnly Then
       ctlWarningPopup.ShowWarning ctlBayNr(Index).hwnd, LoadResString(sInvalidBayNumber), 2000
       If g_iDebug = 13 Then
             Call InfoLog("frmPDD ctlBayNr_Validate - End - GateBay, arrive tractor only - Index: " & Index)
       End If
       Exit Sub
    End If
 
    If ((ctlBayNr.Item(Index).Text = GATE_BAY_TYPE And Len(grdBayDetails.Item(Index).Columns(GRD1_COL_TRAILER).Text) = 0 And Not bDuplicateLoad) _
    Or (ctlBayNr.Item(Index).Text = GATE_BAY_TYPE And Len(grdBayDetails.Item(Index).Columns(GRD1_COL_TRAILER).Text) > 0 And Not bDuplicateLoad And Not bBalNrValidated(Index))) Then
        ' We ONLY allow entry of GATE_BAY_TYPE
        ' if invoked by the ARRIVAL Process
        bLostFocus = True
        bResult = pInst.GateBayProcess(Index + 1)
        If bResult = False Then
            ctlBayNr.Item(Index).Text = sCurBayNum
            Cancel = True
            bBayCompeleted(Index) = False
            bBalNrValidated(Index) = False
            If g_iDebug = 13 Then
                  Call InfoLog("frmPDD ctlBayNr_Validate - End - GateBay, result false - Index: " & Index)
            End If
            Exit Sub
        End If
        If bFromGTForm(Index) Then   '5/17 QS
           bBalNrValidated(Index) = True
           If pInst.PDInfo.hPDHll.Count > 1 Then
              If CheckDuplicate(Index) Then
                 Cancel = True
                 ctlBayNr.Item(Index).Visible = True
                 ctlBayNr.Item(Index).Enabled = True
                 ctlBayNr.Item(Index).SetFocus
                 bBayCompeleted(Index) = False
                 bBalNrValidated(Index) = False
                 If g_iDebug = 13 Then
                      Call InfoLog("frmPDD ctlBayNr_Validate - End - bFromGtForm - Index: " & Index)
                 End If
                 Exit Sub
              Else
                 bBalNrValidated(Index) = True
              End If
           End If
           bctlBayValidateComplete(Index) = True
           If g_iDebug = 13 Then
                      Call InfoLog("frmPDD ctlBayNr_Validate - End - bFromGtForm 2 - Index: " & Index)
                 End If
           Exit Sub
       End If
    ElseIf (Not IsBlank(sCurBayNum)) Or grdBayDetails.Item(Index).Columns(GRD1_COL_PD).Value = HFCS_INDICATOR_TRUE Then  ' Get the Bay info
      If Not bFromF4F5 And Not bDatafromMatchingScreen(Index) And Not bDataFromEloc(Index) Then
        Set oBay = oYard.BayGetSingle(sCurBayNum, HFCS_INDICATOR_FALSE, 13, oClsDb.DBClass)
        If oBay Is Nothing Then
            If Len(Trim$(sCurBayNum)) > 0 Then
                bResult = False
            ElseIf Len(Trim$(sCurBayNum)) = 0 And Not IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Then
                bResult = True
            End If
        Else
            bResult = True
            pInst.bPreDispatched(Index + 1) = False
 
            iBayPosition = 0
            For Each oColBay In g_colBays
               bBayStillExist = False
               For i = 0 To ctlBayNr.UBound
                    If ctlBayNr(i).Text = oColBay.Number Then
                        Exit For
                    End If
               Next i
 
               If i > ctlBayNr.UBound Then
                  g_colBays.Remove iBayPosition + 1
                  bBayStillExist = True
               End If
 
               If Not bBayStillExist Then
                   If oColBay.Number <> oBay.Number Then
                       Set rsBays = oClsDb.CheckForLockedBay(oColBay.Number)
 
                       'this has to be done because the Yard Object unlocks the bay when setting the bay object to another bay
                       'therefore, lock other bays that show if it is not locked.
                       If IsNull(rsBays.Fields("bay_lck_ts").Value) Then
                          'check to see if bay is locked
                          oColBay.LockBay BAY_PRIMARY_LOCK
                       End If
                    End If
                    iBayPosition = iBayPosition + 1
                End If
            Next
 
            iRc = pInst.ProcessBayInfo(oBay)
           
            'If ProcessBayInfo is not successful then we set focus back on Bay Number and exit the sub
            'This logic is hit if the bay is empty and returns no value. At this point there is not need to perform
            'further valiation - JLW8PGC
            If iRc <> DB_SUCCESS Then
                bResult = False
                Cancel = True
                sPreBayNum(Index) = ctlBayNr.Item(Index).Text
                ctlBayNr.Item(Index).SelStart = 0
                ctlBayNr.Item(Index).SelLength = Len(ctlBayNr.Item(Index).Text)
                If g_iDebug = 13 Then
                    Call InfoLog("frmPDD ctlBayNr_Validate - End - If iRc <> DB_SUCCESS (Empty Bay) - Index: " & Index)
                End If
                Exit Sub
            Else
               'Get the info for the multiple loads.
               If pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).szMultLdIr = IR_TRUE Then
                  sServiceTypeCD = oClsDb.Get_Service_Type(Trim$(pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).lCurFopTlrEntity))
               End If
 
               'Check if the loads has been depart?
               For i = 1 To pInst.PDInfo.hPDHll.Count
                  If pInst.PDInfo.hPDHll.Count >= Index + 1 Then   '8/12/05
                      DisplayLoadInfo pInst.PDInfo.hPDHll.Item(Index + 1), Index
                      pInst.raLockedBay(Index + 1) = ctlBayNr.Item(Index).Text
                      bShiftLdInfo = False
                      If i = Index + 1 Then
                         bLostFocus = True
                      End If
                  End If
               Next i
               pInst.bPreDispatched(Index + 1) = True
            End If
        End If
      Else
        If pInst.PDInfo.hPDHll.Count >= Index + 1 Then
           bResult = True ' Data is ready
        End If
        For Each oColBay In g_colBays
           bBayStillExist = False
           For i = 0 To ctlBayNr.UBound
                If ctlBayNr(i).Text = oColBay.Number Then
                    Exit For
                End If
           Next i
 
           If i > ctlBayNr.UBound Then
              g_colBays.Remove iBayPosition + 1
              bBayStillExist = True
           End If
 
           If Not bBayStillExist Then
               If oColBay.Number <> oBay.Number Then
                   Set rsBays = oClsDb.CheckForLockedBay(oColBay.Number)
 
                   'this has to be done because the Yard Object unlocks the bay when setting the bay object to another bay
                   'therefore, lock other bays that show if it is not locked.
                   If IsNull(rsBays.Fields("bay_lck_ts").Value) Then
                      'check to see if bay is locked
                      oColBay.LockBay BAY_PRIMARY_LOCK
                   End If
                End If
 
                iBayPosition = iBayPosition + 1
            End If
        Next
        If pInst.PDInfo.hPDHll.Count < Index + 1 Then
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD ctlBayNr_Validate - End - pInst.PDInfo.hPDHll.Count < Index + 1 - Index: " & Index)
            End If
            Cancel = True
            Exit Sub
        End If
        If Len(Trim$(sCurBayNum)) > 0 Then
            Set colEquip = oClsDb.GetTrailerInfo(sCurBayNum)
           
            If Not colEquip Is Nothing Then
                For Each oTlr In colEquip
                    If oTlr.FOPEntityKey = pInst.PDInfo.hPDHll.Item(Index + 1).lCurFopTlrEntity Or _
                            pInst.PDInfo.hPDHll.Item(Index + 1).lCurFopTlrEntity = 0 Then
                        Exit For
                    End If
                Next
                pInst.PDInfo.hPDHll.Item(Index + 1).szTlrPosCd = oTlr.TrailerPosition
            End If
            Set colEquip = Nothing
            Set oTlr = Nothing
        End If
      End If
    End If
 
'    'Check duplicate loads
    If CheckDuplicate(Index) Then
       Cancel = True
       sPreBayNum(Index) = ctlBayNr.Item(Index).Text
    End If
  If Cancel = False Then
 
    If bResult Then
        If pInst.PDInfo.hPDHll.Count > 0 Then   '8/12/05
           bResult = pInst.CheckPDLoads()
        End If
        If bResult = False Then
            Cancel = True
            If Len(ctlBayNr.Item(Index).Text) = 0 Then
                ProcessLockBay Format$(pInst.raLockedBay(Index + 1), "0000"), Nothing, BAY_UNLOCK
                pInst.raLockedBay(Index + 1) = 0
            End If
            sPreBayNum(Index) = ctlBayNr.Item(Index).Text
            bLostFocus = False
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD ctlBayNr_Validate - End - not bResult CheckPDLoads - Index: " & Index)
            End If
            Exit Sub
        End If
    End If
   
    If Len(ctlBayNr.Item(Index).Text) > 0 And Not IsUserShiftTabbing Then
        SetColumnLocksForEmptyLoad (Index)
        Call SetGridForEdit(Index)
        Call ReTab(LIST_GRIDS) ' "LIST_Grids")
        If pInst.PDInfo.hPDHll.Count >= Index + 1 Then
            If pInst.PDInfo.hPDHll(Index + 1).SameLoadName = True And pInst.PDInfo.hPDHll(Index + 1).szLdCd <> "E" Then
                txtDolly(Index).Enabled = True
            ElseIf pInst.PDInfo.hPDHll(Index + 1).szLdCd <> "E" And (Not IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) And pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID <> pInst.PDInfo.SegmentSystemNumberOID) Then
                grdBayDetails(Index).Split = 3
                grdBayDetails(Index).Col = GRD1_COL_RI
            Else
                If pInst.PDInfo.hPDHll(Index + 1).szLdCd = "E" Then
                'Switched order of next two lines to match best practice and correct Defect 486 - JLW8PGC
                    grdBayDetails(Index).Split = 2
                    grdBayDetails(Index).Col = GRD1_COL_DEST
                Else
                    grdBayDetails(Index).Split = 3
                    grdBayDetails(Index).Col = GRD1_COL_RI
                   
                End If
            End If
        End If
    End If
   
 
    'THIS IS YES   QS 2/3/05
 
    If bResult Then
        If (pInst.PDInfo.hPDHll(Index + 1).szLdCd = "E" And (IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Or pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID <> pInst.PDInfo.SegmentSystemNumberOID)) And Not IsUserShiftTabbing Then
             bPredispatchEmptyWithTrailer(Index) = True
             SetColumnLocksForEmptyLoad (Index)
             Call SetGridForEdit(Index)
             Call ReTab(LIST_GRIDS)
             If (Len(Trim$(ctlBayNr(Index).Text)) > 0) Then
               '  grdBayDetails(Index).SetFocus
                 'Switched order of next two lines to match best practice and correct Defect 486; Split the Col Focus - JLW8PGC
                 grdBayDetails(Index).Split = 2
                 grdBayDetails(Index).Col = GRD1_COL_DEST
             End If
             bLostFocus = True
        ElseIf pInst.PDInfo.hPDHll(Index + 1).SameLoadName = True Then
            txtDolly(Index).Enabled = True
            txtDolly(Index).SetFocus
        End If
   
    ElseIf Not bResult And grdBayDetails.Item(Index).Columns(GRD1_COL_PD).Value = HFCS_INDICATOR_TRUE Or Len(Trim$(ctlBayNr(Index).Text)) > 0 Then
        Cancel = True
        sPreBayNum(Index) = ctlBayNr.Item(Index).Text
        ctlBayNr.Item(Index).SelStart = 0
        ctlBayNr.Item(Index).SelLength = Len(ctlBayNr.Item(Index).Text)
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_Validate end bResult false.")
        End If
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ctlBayNr_Validate - End -  ElseIf Not bResult And... - Index: " & Index)
        End If
        Exit Sub
    End If
 
    End If
 
    'To find the dolly which is on the bay
    If Not Cancel Then
        If pInst.PDInfo.hPDHll.Count >= Index + 1 And Not bFromGTForm(Index) Then
           sDollyOnBay = Normalize(CStr(oClsDb.GetDollyOnBay(ctlBayNr.Item(Index).Text, pInst.PDInfo.hPDHll.Item(Index + 1).szTlrPosCd)))
           If Len(sDollyOnBay) = 0 Then
          '    txtDolly.Item(Index).Text = ""
           Else   'Dolly attached to this bay.
              iRc = pdMsgBox(LoadResString(sDollyattatched) & vbCr & LoadResString(sDoYouWanttoUseIt), vbQuestion + vbYesNo + vbDefaultButton2, LoadResString(sQuittingPredispatchWARNING), Me.hwnd)
              If iRc = vbYes Then
                  txtDolly.Item(Index).Text = sDollyOnBay
              Else
                  txtDolly.Item(Index).Text = vbNullString
                  pInst.PDInfo.hPDHll.Item(Index + 1).lCurFopDolEntity = 0
             End If
           End If
        End If
    End If
    If Cancel Then
        ' the following locks and unlocks a bay
        If pInst.raLockedBay(Index + 1) <> 0 Then
           'Unlock the Bay
           pInst.raLockedBay(Index + 1) = 0
        End If
        ProcessLockBay sCurBayNum, Nothing, BAY_UNLOCK
        sLastBayNr(Index + 1) = ""
        ctlBayNr.Item(Index).Visible = True
        ctlBayNr.Item(Index).Enabled = True
        ctlBayNr.Item(Index).SetFocus
        ctlBayNr.Item(Index).SelectText
        pInst.bPreDispatched(Index + 1) = False
    Else
        pInst.bPreDispatched(Index + 1) = True
        bBalNrValidated(Index) = True
        Call SetF2F7Button
        bDataFromEloc(Index) = False
    End If
 
    If pInst.PDInfo.hPDHll.Count >= Index + 1 Then
       If (pInst.PDInfo.hPDHll(Index + 1).szLdCd = LD_CD_EMPTY And pInst.PDInfo.hPDHll(Index + 1).szDestinSrt = LD_CD_EMPTY) And Not bFromGTForm(Index) And Not bPredispatchEmpty(Index) And Not bPredispatchEmptyWithTrailer(Index) Then
          Set rsRecordset = oClsDb.GetScheduleNames(pInst.PDInfo.hPDHll(Index + 1))
          If Not rsRecordset Is Nothing Then
              If rsRecordset.RecordCount > 0 Then
                  rsRecordset.Find "LD_SEQ_NR = '" & Format$(pInst.PDInfo.hPDHll(Index + 1).iSequenceNr, "00") & "' "
                  If rsRecordset.EOF Then
                    rsRecordset.MoveFirst
                  ElseIf IsNull(rsRecordset.Fields("ASN_TLR_NR").Value) Then
                    Do While Not rsRecordset.EOF
                        If Not IsNull(rsRecordset.Fields("ASN_TLR_NR").Value) Then
                            If Len(Trim$(rsRecordset.Fields("ASN_TLR_NR").Value)) > 0 Then
                                Exit Do
                            End If
                        End If
                        rsRecordset.MoveNext
                        rsRecordset.Find "LD_SEQ_NR = '" & Format$(pInst.PDInfo.hPDHll(Index + 1).iSequenceNr, "00") & "' "
                    Loop
                    If rsRecordset.EOF Then rsRecordset.MoveFirst
                  End If
                  lOrgFRDKey(Index) = rsRecordset.Fields("FRD_TLR_EQP_GEN_NR").Value
              End If
          End If
          Call oClsDb.DBClass.CloseRecordSet(rsRecordset)
       ElseIf bDatafromMatchingScreen(Index) = False Then
          lOrgFRDKey(Index) = pInst.PDInfo.hPDHll.Item(Index + 1).lCurForTlrEntity
          bMatchingLoadFound(Index) = False
       End If
    End If
    sPreBayNum(Index) = ctlBayNr.Item(Index).Text
   
    If Not Cancel Then
        bctlBayValidateComplete(Index) = True
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlBayNr_Validate - End - Index: " & Index)
    End If
End If
'bRoutingCanceled = False
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "ctlBayNr_Validate"
    Call oProc.update_error_object(Me, gsError)
    Call oClsDb.DBClass.CloseRecordSet(rsRecordset)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
    End If
End Sub
 
Private Sub ctlWkEndingDate_LostFocus()
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlWkEndingDate_LostFocus - Begin")
    End If
    If ctlWkendingDate.Validate = False Then
        If ctlWkendingDate.Enabled = False Then ctlWkendingDate.Enabled = True
        ctlWkendingDate.SetFocus
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ctlWkEndingDate_LostFocus - End")
    End If
End Sub
 
Private Sub drpValTrailerType_DropDownClose(Index As Integer)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD drpValTrailerType_DropDownClose - Begin - Index: " & Index)
    End If
    m_bDropDownTlrTyp(Index) = False
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD drpValTrailerType_DropDownClose - End - Index: " & Index)
    End If
End Sub
 
Private Sub drpValTrailerType_DropDownOpen(Index As Integer)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD drpValTrailerType_DropDownOpen - Begin - Index: " & Index)
    End If
    drpValTrailerType(Index).Columns(0).Width = 500
    m_bDropDownTlrTyp(Index) = True
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD drpValTrailerType_DropDownOpen - End - Index: " & Index)
    End If
End Sub
 
Private Sub Form_KeyPress(KeyAscii As Integer)
'==============================================================
' Description:  Process operations for the form key press event
'==============================================================
If g_iDebug = 13 Then
    Call InfoLog("frmPDD Form_KeyPress - Begin")
End If
  On Error GoTo FormKeyPressError
 
 If gbLoadPreDispatched = True Or m_bInitializing Then
     KeyAscii = 0
     If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_KeyPress - End - If gbLoadPreDispatched = True")
     End If
     Exit Sub
End If
    
 oKeyDefs.CheckForEnterKey Me.hwnd
 
    
 If bBayCompeleted(iBayIndex) And Me.ActiveControl Is ctlBayNr.Item(iBayIndex) Then
    If KeyAscii = oKeyDefs.Enter_Key Then
       EnterAsTab KeyAscii
    Else
       KeyAscii = 0
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD Form_KeyPress key not enter key - End")
        End If
       Exit Sub
    End If
  End If
 
 If Me.ActiveControl Is cboDow Then
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD Form_KeyPress DOW Code active - End")
    End If
   Exit Sub
End If
  If pInst.bArriveFlag And _
     bBayCompeleted(iBayIndex) And Me.ActiveControl Is ctlBayNr.Item(iBayIndex) Then
     KeyAscii = 0
     If g_iDebug = 13 Then
       Call InfoLog("frmPDD Form_KeyPress arrive flag - End")
     End If
     Exit Sub
  End If
 
  If (KeyAscii = oKeyDefs.Enter_Key) Then
       EnterAsTab KeyAscii
       KeyAscii = KeyToUpperCase(KeyAscii)
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD Form_KeyPress Enter key - End")
        End If
       Exit Sub
    End If
 
  
  ' Convert to uppercase
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
 
  ' Prevent beeping problem by setting KeyAscii to 0 on <ESC>
  If oKeyDefs Is Nothing Then
    Set oKeyDefs = New GLOBALDEFS.clsKeyCodes
  End If
  If (KeyAscii = oKeyDefs.Execute_key) Or (KeyAscii = oKeyDefs.Keypad_Execute_Key) Then
    SendKeys "~"
    KeyAscii = 0
  End If
 
  ' Prevent nonalphabetic characters from appearing
  If IsNonAlphaNumeric(Chr$(KeyAscii)) Then
    If Not KeyAscii = Asc("/") Then
      'Team Track #434 - allow user to enter . in eloc field when ELOC is TRM.)
      If Not ((KeyAscii = Asc(".") And Me.ActiveControl Is txtEloc And Mid$(txtEloc.Text, 1, 3) = "TRM") Or (KeyAscii = Asc("*") And Me.ActiveControl.Name = "grdBayDetails")) Then
          KeyAscii = 0
      End If
    End If
  End If
 
  ' Do not accept problem characters
  If Chr$(KeyAscii) = "'" Then
    KeyAscii = 0
  End If
     
  ' Prevent beeping problem by setting KeyAscii to 0 on <ESC>, <CR>
  If KeyAscii = PLUS_KEY Then
    Form_KeyDown vbKeyF10, 0
    KeyAscii = 0
  End If
  
 If g_iDebug = 13 Then
    Call InfoLog("frmPDD Form_KeyPress - End")
End If
Exit Sub
 
FormKeyPressError:
  Set oEventlog = New HFErrorObject.clsEventLogs
  Call oErrorObject.error_routine(oEventlog.FeederShell, Err.Number, Err.Description, oProcessInfo, ERROR_MSG, FEEDER_DISPATCH_DRIVER)
  Set oEventlog = Nothing
  Err.Clear
End Sub
 
Private Sub Form_Load()
Dim i As Integer
 
On Error GoTo ErrorExit
  
    Set oProc = New HFErrorObject.clsProcessInformation
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_Load() - Begin")
    End If
   
    m_bInitializing = True
  '  m_bGridGotFocus = False
    bTrailerError = False
    bLoadRoutingError = False
    bEscPressedOnError = False
   
    bPredispatchEmpty(0) = False
    bPredispatchEmpty(1) = False
    bPredispatchEmpty(2) = False
   
    bPredispatchEmptyWithTrailer(0) = False
    bPredispatchEmptyWithTrailer(1) = False
    bPredispatchEmptyWithTrailer(2) = False
    bSetFocusonBay = False
   ' bGridGotFocus = False
   ' bRoutingCanceled = False
    m_bClickCancelPD = False
   
    'Call InfoLog("frmPDD Begin Form_Load", WARNING_MSG)
    Call InitialVaribles
   
    ' Determine local hub and SLIC
    sLocalCn = Left$(oRegistry.LocalHub, 2)
    sLocalHub = oRegistry.LocalHub
    g_iDebug = oRegistry.ArrivalDebug
   
    Call SetupGrid
    ' TT000394
    'Dim myList As String
    'myList = "List_No_Grids"
    ReTab (LIST_NO_GRID) ' (myList)
   
    grdBayDetails(0).Bookmark = Null
    grdBayDetails(1).Bookmark = Null
    grdBayDetails(2).Bookmark = Null
   
    With grdBayDetails(0).Columns(GRD1_COL_PD).ValueItems
        .Translate = True
        .Presentation = dbgCheckBox
    End With
    With grdBayDetails(1).Columns(GRD1_COL_PD).ValueItems
        .Translate = True
        .Presentation = dbgCheckBox
    End With
    With grdBayDetails(2).Columns(GRD1_COL_PD).ValueItems
        .Translate = True
        .Presentation = dbgCheckBox
    End With
    grdBayDetails(0).ReBind
    grdBayDetails(1).ReBind
    grdBayDetails(2).ReBind
   
    gbLoadPreDispatched = False
    g_bRoutingOpen = False
    bPreDispatched = False
    mbClearData = False
   
'dsb ~~~ the next statement must be changed to
'                PreDispatchDriver
'         requires change to the owsinfo object.
    Call oWSInfo.Register_Application( _
                              OGlobalFormNames.PredipatchDriver, _
                              App.FileDescription, _
                              hwnd)
    Screen.MousePointer = vbHourglass
    SetScreenDetails
    SetupToolbar
    ' set dow and week ending dates, list box
    Call setDowAndWkEndingDate
    InitPdInfo
    'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
    Screen.MousePointer = vbNormal
 
    gbExitScreen = False
    m_bSettingLoadInfo = False
   
    'Set the Application Wide Help File
    App.HelpFile = oRegistry.HelpPath
 
    ctlRecordCount.Text = ""
    mbElocValid = False
   
   
    
     Set rsValidTrailerTypesDropDown = oClsDb.Get_Valid_Trailer_Type_Codes
     drpValTrailerType(0).DataSource = rsValidTrailerTypesDropDown
     drpValTrailerType(1).DataSource = rsValidTrailerTypesDropDown
     drpValTrailerType(2).DataSource = rsValidTrailerTypesDropDown
     Call drpValTrailerType(0).Move(100, 100, 1000, 1000)
     Call drpValTrailerType(1).Move(100, 100, 1000, 1000)
     Call drpValTrailerType(2).Move(100, 100, 1000, 1000)
   
    If pInst.bArriveFlag Then
       If Not (Len(Trim$(txtDriverName.Text)) = 0 And (bFromGtF8 Or bArrivalTractorOnly Or gbFromArrival)) Then
        grdMultLegView.Enabled = True
'        txtEloc.Enabled = True
        SetAvailMenuAndTool MAT_BASE_DATA0, True, False
        txtJobNr.Enabled = False
        txtDriverName.Enabled = False
        ctlWkendingDate.Enabled = False
        cboDow.Enabled = False
       Else
        txtJobNr.Enabled = True
        txtJobNr.SetFocus
       End If
    Else
       Show
       txtJobNr.Enabled = True
       txtJobNr.SetFocus
    End If
   
    g_bShowMatchingLoad = True
    m_bInitializing = False
If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_Load() - End")
End If
' Err.Raise 7, TypeName(Me), "Robert Clark"
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "Form Load"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                 oProc, _
                                 ERROR_MSG, _
                                 FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
        m_bInitializing = False
     '   Unload Me
    End If
End Sub
'<WR01024><ADID: dmx7plm> - Start
Private Sub Form_LostFocus()
On Error Resume Next
  Dim ctl As Control
  If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_LostFocus() - Begin")
  End If
 
    For Each ctl In Me.Controls
        If ctl.Name = Me.Tag Then
            If ctl.Name = "ctlBayNr" Or ctl.Name = "txtDolly" Or ctl.Name = "grdBayDetails" Then
                If ctl.Item(ictlIndex).Enabled = False Then ctl.Item(ictlIndex).Enabled = True
                ctl.Item(ictlIndex).SetFocus
                Exit For
             Else
                If ctl.Enabled = False Then ctl.Enabled = True
                ctl.SetFocus
                Exit For
             End If
        End If
    Next
 
If Me.Tag = "grdBayDetails" And ctlControls.Name = Me.Tag Then
    grdBayDetails(ictlIndex).Enabled = True
    grdBayDetails(ictlIndex).SetFocus
End If
 
If Me.Tag = "ctlBayNr" And ctlControls.Name = Me.Tag Then
    ctlBayNr(ictlIndex).Visible = True
    ctlBayNr(ictlIndex).Enabled = True
    ctlBayNr(ictlIndex).SetFocus
End If
 
If Me.Tag = "txtDolly" And ctlControls.Name = Me.Tag Then
    txtDolly(ictlIndex).Enabled = True
    txtDolly(ictlIndex).SetFocus
End If
If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_LostFocus() - End")
End If
End Sub
'<WR01024><ADID: dmx7plm> - End
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim tMsg As MSG
Dim iRc As Integer
 
    On Error Resume Next
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_QueryUnload - Begin")
    End If
   'clear keyboard buffer
    Do While PeekMessage(tMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
        'do nothing, just clear out buffer
    Loop
    If bExitScreen Then
       'Call InfoLog("frmPDD End Form_QueryUnload, Exit Screen", WARNING_MSG)
       If g_iDebug = 13 Then
            Call InfoLog("frmPDD Form_QueryUnload - End - If bExitScreen")
        End If
       Exit Sub
    End If
    If UnloadMode = vbFormCode Or UnloadMode = vbFormControlMenu Then
        If Not bPreDispatched Then
            iRc = pdMsgBox(LoadResString(sDoYouWishExitPredispatch) & vbCr & LoadResString(sWithoutPredispatchingLoads), vbQuestion + vbYesNo + vbDefaultButton2, LoadResString(sQuittingPredispatchWARNING))
 
            If iRc = vbYes Then
                Cancel = 0
            Else
                Cancel = -1
            End If
        End If
    End If
    'Call InfoLog("frmPDD End Form_QueryUnload", WARNING_MSG)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_QueryUnload - End")
    End If
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim icon As Integer
Dim vntTemp As Variant
Dim tMsg As MSG
Dim j As Integer
 
    On Error Resume Next
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD Form_Unload - Begin")
    End If
       
    '<WR01024><ADID: dmx7plm> - Start
    If ClosefrmPDD = False Then
        Cancel = True
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD Form_Unload - End - If ClosefrmPDD = False")
        End If
        Exit Sub
    End If
    '<WR01024><ADID: dmx7plm> - End
    'clear keyboard buffer
    Do While PeekMessage(tMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
        'do nothing, just clear out buffer
    Loop
   
    'kg - moved before unregister so it can be handled appropriately in Arrive
    'Tractor screen
    'Sent back PDP info to Arrival Tractor Only screen.
    If pInst.bArriveFlag And bArrivalTractorOnly Then
       If bPreDispatched Then
          Call frmKill.PreDispatchDown(pInst.bArriveFlag)  'QS 11/29/04
       Else
          Call frmKill.PreDispatchDown(IR_FALSE)
       End If
    End If
   
'    oWSInfo.UnRegister_Application OGlobalFormNames.PredipatchDriver
'    Set oWSInfo = Nothing
    Set OGlobalFormNames = Nothing
    For i = 1 To 3
        If pInst.raLockedBay(i) <> 0 Then
            ' Unlock the Bay after PreDispatch
            ProcessLockBay ctlBayNr(i - 1).Text, Nothing, BAY_UNLOCK
            pInst.raLockedBay(i) = 0
        End If
        Set oColLoads(i - 1) = Nothing
        For j = 0 To 2
            Set oOldColLoads(i - 1, j) = Nothing
        Next j
    Next
    Set oProc = New HFErrorObject.clsProcessInformation
    gsError = "Form_Unload"
    Call oProc.update_error_object(Me, gsError)
   
    
'    For icon = 1 To 50
'      'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
'    Next icon
   
    mbElocValid = False
    mbClearData = False
    Set rsRecordset = Nothing
   
    For i = 1 To 3
        sLastBayNr(i) = ""
    Next i
   
    txtLastDriverName = ""
   
    
     'Sent back PDP info to Arrival screen
    If pInst.bArriveFlag And Not bArrivalTractorOnly And bFromGtF8 Then
       If oLDInfo Is Nothing Then
          Set oLDInfo = New Collection
       End If
       'Add load info
       If Not oLoadInfo(0) Is Nothing Then
           oLDInfo.Add oLoadInfo(0)
       End If
       If Not oLoadInfo(1) Is Nothing Then
          oLDInfo.Add oLoadInfo(1)
       End If
       If Not oLoadInfo(2) Is Nothing Then
          oLDInfo.Add oLoadInfo(2)
       End If
      
       'Added Dolly # in collection
       sDolly = Me.txtDolly(0).Text & "|" & Me.txtDolly(1).Text & "|" & Me.txtDolly(2).Text
       If iPredispatch = 0 Then
            PDPInfo1 = vbNullString
            PDPInfo2 = vbNullString
            sRouteCodeMessage(0) = vbNullString
            sRouteCodeMessage(1) = vbNullString
            sRouteCodeMessage(2) = vbNullString
            sRouteCodeChanged(0) = vbNullString
            sRouteCodeChanged(1) = vbNullString
            sRouteCodeChanged(2) = vbNullString
       End If
       vntTemp = Array(sDolly, sPDPinfo, oLoadInfo, PDPInfo1, PDPInfo2, sRouteCodeMessage, sRouteCodeChanged)
       Call frmKill.PreDispatchDone(vntTemp)
    End If
 
  
    sPDPinfo = vbNullString
    sDolly = vbNullString
    Set oLoadInfo(0) = Nothing
    Set oLoadInfo(1) = Nothing
    Set oLoadInfo(2) = Nothing
    Set oLDInfo = Nothing
   
    gbFromArrival = False
    bArrivalTractorOnly = False
    Set oRegistry = Nothing
'    Set oWSInfo = Nothing
'    Set oWSNotify = Nothing
'    Set OGlobalFormNames = Nothing
    Set oGlobalMsgs = Nothing
    Set oColHeads = Nothing
    Set oHelp = Nothing
    Set oErrorObject = Nothing
    Set oKeyDefs = Nothing
    Set oNotify = Nothing
    Set pInst = Nothing
    Set gpData = Nothing   'QS add it to clear the grid after close the form.
    Set oYard = Nothing
    Set oBay = Nothing
    Call oClsDb.DBClass.CloseRecordSet(rsValidTrailerTypesDropDown)
 
    'Remove the objects from the collection
    If g_colBays.Count > 0 Then
      For i = g_colBays.Count To 1 Step -1
        g_colBays.Remove (i)
      Next i
    End If
   
    Set frmKill = Nothing
  
    'Clear the array that contains the arrival data.
    For i = 1 To 3
      With raArrivalData(i)
          .dtDueDt = "12:00:00 AM"
          .dtLdCrtDt = "12:00:00 AM"
          .dtSchActDt = "12:00:00 AM"
          .iActHazMatQy = 0
          .iActPkgPr = 0
          .lActLdMsgId = 0
          .lActPkgQy = 0
          .lActTlrMsgId = 0
          .lFopDolEntKey = 0
          .lFopTlrEntKey = 0
          .lFopVehEntKey = 0
          .lForTlrEntKey = 0
          .sActDolNr = ""
          .sActDvrNa = ""
          .sActFdrJobNa = ""
          .sActLdCd = ""
          .sActOrgSrtTypCd = ""
          .sActQyPrCd = ""
          .sActRteTypCd = ""
          .sActTlrNr = ""
          .sActTlrTypCd = ""
          .sActTrcNr = ""
          .sBayNr = ""
          .sFdrJobDmcCnyCd = ""
          .sFdrJobDmcSAB = ""
          .sLdDtnCnyCd = ""
          .sLdDtnSlcAbrNa = ""
          .sLdDtnSrtTypCd = ""
          .sLdOrgCnyCd = ""
          .sLdOrgSlcAbrNa = ""
          .sLdSeqNr = ""
          .sRemarks = ""
      End With
      Next i
    oWSInfo.UnRegister_Application OGlobalFormNames.PredipatchDriver
    Set oWSInfo = Nothing
    Set OGlobalFormNames = Nothing
      'Call InfoLog("frmPDD Form_Unload end", WARNING_MSG)
      If g_iDebug = 13 Then
           Call InfoLog("frmPDD Form_Unload - End")
      End If
End Sub
 
Private Sub grdBayDetails_BeforeRowColChange(Index As Integer, Cancel As Integer)
  Dim iReturnMessage As Integer
  Dim sCurrentRouteCode As String
  Dim sNewRouteCode As String
  Dim bInvalidLoadRouting As Boolean
 
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdBayDetails_BeforeRowColChange - Begin - Index: " & Index)
    End If
 
  If grdBayDetails(Index).Col = -1 Then
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdBayDetails_BeforeRowColChange - End - If grdBayDetails(Index).Col = -1 - Index: " & Index)
    End If
    Exit Sub
  End If
   
  'If there is a warning popped up and the user has not
  'attempted to correct then they cannot change rows or cells.
  If ctlWarningPopup.WarningVisible Then
    ctlWarningPopup.ClearWarning
  End If
   
  If grdBayDetails(Index).Col = GRD1_COL_DEST Or _
         grdBayDetails(Index).Col = GRD1_COL_DS Or _
         grdBayDetails(Index).Col = GRD1_COL_RI Or _
         grdBayDetails(Index).Col = GRD1_COL_TYPE Then
       
        If IsNull(pInst.PDInfo.hPDHll(Index + 1).szRtId) Then
            sCurrentRouteCode = vbNullString
        Else
            sCurrentRouteCode = Trim$(pInst.PDInfo.hPDHll(Index + 1).szRtId)
        End If
       
        Cancel = IIf(PerformFieldLevelValidation(Index), False, True)
       
        If Cancel = True Then
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD grdBayDetails_BeforeRowColChange - End - If Cancel = True - Index: " & Index)
            End If
            Exit Sub
        ElseIf bPredispatchEmptyWithTrailer(Index) = True Then
           If grdBayDetails(Index).Col = GRD1_COL_RI Then
               If IsNull(grdBayDetails(Index).Columns(GRD1_COL_RI).Value) Then
                   sNewRouteCode = vbNullString
               Else
                   sNewRouteCode = Trim$(grdBayDetails(Index).Columns(GRD1_COL_RI).Value)
               End If
               If sCurrentRouteCode <> sNewRouteCode Then
                   pInst.PDInfo.hPDHll(Index + 1).SendPredispatchPreforecast = True
               End If
               pInst.PDInfo.hPDHll(Index + 1).szRtId = sNewRouteCode
               If bSaveData = False Then
                   SetAvailMenuAndTool mat_bay, True, False
                   bBayCompeleted(Index) = True
                   txtDolly(Index).Enabled = True
                   txtDolly(Index).SetFocus
                   If Index = 2 Then bSaveData = True
               End If
           ElseIf grdBayDetails(Index).Col = GRD1_COL_DS Then
               Call pInst.UpdatePredispatchWithEmpty(Index)
               grdBayDetails(Index).Columns(GRD1_COL_SEQ).Value = "00" 'oClsDb.GetEmptySequenceNumber(pInst.PDInfo, Index + 1)  'Format$(pInst.PDInfo.hPDHll(Index + 1).iSequenceNr, "00")
               pInst.PDInfo.hPDHll(Index + 1).iSequenceNr = CInt(grdBayDetails(Index).Columns(GRD1_COL_SEQ).Value)
           ElseIf grdBayDetails(Index).Col = GRD1_COL_DEST Then
               grdBayDetails(Index).Columns(GRD1_COL_OS).Value = LD_CD_EMPTY
               pInst.PDInfo.hPDHll(Index + 1).szOriginSrt = LD_CD_EMPTY
           End If
        ElseIf bPredispatchEmpty(Index) = True Then
           If grdBayDetails(Index).Col = GRD1_COL_DS Then
               grdBayDetails(Index).Columns(GRD1_COL_SEQ).Value = "00" 'oClsDb.GetEmptySequenceNumber(pInst.PDInfo, Index + 1)  'Format$(pInst.PDInfo.hPDHll(Index + 1).iSequenceNr, "00")
               If pInst.PDInfo.hPDHll(Index + 1).iSequenceNr > 0 Then
                   pInst.PDInfo.hPDHll.Item(Index + 1).lCurForTlrEntity = 0
               End If
               pInst.PDInfo.hPDHll(Index + 1).iSequenceNr = CInt(grdBayDetails(Index).Columns(GRD1_COL_SEQ).Value)
           ElseIf grdBayDetails(Index).Col = GRD1_COL_RI Then
               If IsNull(grdBayDetails(Index).Columns(GRD1_COL_RI).Value) Then
                   sNewRouteCode = vbNullString
               Else
                   sNewRouteCode = Trim$(grdBayDetails(Index).Columns(GRD1_COL_RI).Value)
               End If
               If sCurrentRouteCode <> sNewRouteCode Then
                    pInst.PDInfo.hPDHll(Index + 1).EmptyRouteCodeChanged = True
               End If
               If Index = 1 Then
                If pInst.PDInfo.hPDHll(1).NextSameLoadName = True And pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True Then
                    pInst.PDInfo.hPDHll(1).szRtId = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    grdBayDetails(0).Columns(GRD1_COL_RI).Value = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False
                    pInst.PDInfo.hPDHll(1).SendPredispatchPreforecast = True
                End If
               End If
            If Index = 2 Then
                If pInst.PDInfo.hPDHll(1).NextSameLoadName = True And pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True Then
                    pInst.PDInfo.hPDHll(1).szRtId = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    grdBayDetails(0).Columns(GRD1_COL_RI).Value = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False
                    pInst.PDInfo.hPDHll(1).SendPredispatchPreforecast = True
                End If
                If pInst.PDInfo.hPDHll(2).NextSameLoadName = True And pInst.PDInfo.hPDHll(2).InvalidLoadRouting = True Then
                    pInst.PDInfo.hPDHll(2).szRtId = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    grdBayDetails(1).Columns(GRD1_COL_RI).Value = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    pInst.PDInfo.hPDHll(2).InvalidLoadRouting = False
                    pInst.PDInfo.hPDHll(2).SendPredispatchPreforecast = True
                End If
            End If
               If bSaveData = False Then
                SetAvailMenuAndTool mat_bay, True, False
                bBayCompeleted(Index) = True
               ' bGridGotFocus = False
                txtDolly(Index).Enabled = True
                txtDolly(Index).SetFocus
                If Index = 2 Then bSaveData = True
            End If
           End If
        ElseIf grdBayDetails(Index).Col = GRD1_COL_RI Then
            If IsNull(grdBayDetails(Index).Columns(GRD1_COL_RI).Value) Then
                sNewRouteCode = vbNullString
            Else
                sNewRouteCode = Trim$(grdBayDetails(Index).Columns(GRD1_COL_RI).Value)
            End If
            If sCurrentRouteCode <> sNewRouteCode Then
                pInst.PDInfo.hPDHll(Index + 1).SendPredispatchPreforecast = True
            End If
            pInst.PDInfo.hPDHll(Index + 1).szRtId = sNewRouteCode
           
            If Index = 1 Then
                If pInst.PDInfo.hPDHll(1).NextSameLoadName = True And pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True Then
                    pInst.PDInfo.hPDHll(1).szRtId = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    grdBayDetails(0).Columns(GRD1_COL_RI).Value = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False
                    pInst.PDInfo.hPDHll(1).SendPredispatchPreforecast = True
                End If
            End If
            If Index = 2 Then
                If pInst.PDInfo.hPDHll(1).NextSameLoadName = True And pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True Then
                    pInst.PDInfo.hPDHll(1).szRtId = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    grdBayDetails(0).Columns(GRD1_COL_RI).Value = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False
                    pInst.PDInfo.hPDHll(1).SendPredispatchPreforecast = True
                End If
                If pInst.PDInfo.hPDHll(2).NextSameLoadName = True And pInst.PDInfo.hPDHll(2).InvalidLoadRouting = True Then
                    pInst.PDInfo.hPDHll(2).szRtId = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    grdBayDetails(1).Columns(GRD1_COL_RI).Value = pInst.PDInfo.hPDHll(Index + 1).szRtId
                    pInst.PDInfo.hPDHll(2).InvalidLoadRouting = False
                    pInst.PDInfo.hPDHll(2).SendPredispatchPreforecast = True
                End If
            End If
            If bSaveData = False Then
                SetAvailMenuAndTool mat_bay, True, False
                bBayCompeleted(Index) = True
              '  bGridGotFocus = False
                txtDolly(Index).Enabled = True
                txtDolly(Index).SetFocus
                If Index = 2 Then bSaveData = True
                   
            End If
        End If
  End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD grdBayDetails_BeforeRowColChange - End - Index: " & Index)
    End If
Exit Sub
 
Error_Handler:
glErrNum = 400
gsError = "grdBayDetails_BeforeRowColChange"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
Cancel = True
If oErrorObject.error_routine(oEventlog.FeederShell, _
                           IIf(Err.Number <> 0, Err.Number, glErrNum), _
                           Err.Description & " Module:" & gsError, _
                           oProc, _
                           ERROR_MSG, _
                           FEEDER_DISPATCH_DRIVER) Then
'Set oEventlog = Nothing
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'Unload Me
End If
End Sub
 
Private Sub grdBayDetails_ClassicRead(Index As Integer, _
                                      Bookmark As Variant, _
                                      ByVal Col As Integer, _
                                      Value As Variant)
 
Dim iCursorLoc As Integer
Dim bShowDolly As Boolean
Dim szDolNr As String    ' dolly number
Dim i As Integer
Dim FDR_HIGHLIGHT_COLOR As Long
Dim CLR_NEUTRAL As Long
   
    On Error GoTo ErrorExit
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdBayDetails_ClassicRead Col = " & Col & "- Begin - Index: " & Index)
    End If
    FDR_HIGHLIGHT_COLOR = vbRed
    CLR_NEUTRAL = vbWindowText
    i = Index
   
    If gpData Is Nothing Then ' Clear grid
        Select Case Col
        Case GRD1_COL_BAY
            ' Bay Number
            ctlBayNr(i).Text = ""
            grdBayDetails(i).Columns(GRD1_COL_BAY).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_BAY).Font.Bold = False
        Case GRD1_COL_TRAILER
            ' Reset font information
            grdBayDetails(i).Columns(GRD1_COL_TRAILER).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_TRAILER).Font.Bold = False
           
        Case GRD1_COL_TYPE
            grdBayDetails(i).Columns(GRD1_COL_TYPE).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_TYPE).Font.Bold = False
           
        Case GRD1_COL_PCS
            ' Pieces-Percent
            grdBayDetails(i).Columns(GRD1_COL_PCS).Font.Underline = False
        Case GRD1_COL_PER
            grdBayDetails(i).Columns(GRD1_COL_PER).Font.Underline = False
       
        Case GRD1_COL_ORIG
            ' Load Name
            grdBayDetails(i).Columns(GRD1_COL_ORIG).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_ORIG).Font.Bold = False
           
        Case GRD1_COL_OS
            grdBayDetails(i).Columns(GRD1_COL_OS).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_OS).Font.Bold = False
           
        Case GRD1_COL_DEST
            grdBayDetails(i).Columns(GRD1_COL_DEST).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_DEST).Font.Bold = False
           
        Case GRD1_COL_DS
            grdBayDetails(i).Columns(GRD1_COL_DS).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_DS).Font.Bold = False
           
        Case GRD1_COL_SEQ
            grdBayDetails(i).Columns(GRD1_COL_SEQ).ForeColor = CLR_NEUTRAL
            grdBayDetails(i).Columns(GRD1_COL_SEQ).Font.Bold = False
   
        Case GRD1_COL_LDCD
            ' Load Code
            grdBayDetails(i).Columns(GRD1_COL_LDCD).ForeColor = CLR_NEUTRAL
        End Select
        Value = ""
    Else ' Populate grid
        With gpData
            Select Case Col
            Case GRD1_COL_BAY
                ' Bay Number
                If ctlBayNr(i).Text <> "" Then
                   Value = ctlBayNr(i).Text
                Else
                   Value = .szBayNr
                End If
               
            Case GRD1_COL_POS
                 ' Position
                Value = .szTlrPosCd
               
            Case GRD1_COL_PD
                If pInst.PDInfo.hPDHll.Count > i Then
                    If Not IsNull(pInst.PDInfo.hPDHll(i + 1).AssignedPDOID) Then
                        If Not IsNull(pInst.PDInfo.SegmentSystemNumberOID) Then
                            If Trim$(pInst.PDInfo.SegmentSystemNumberOID) = Trim$(pInst.PDInfo.hPDHll(i + 1).AssignedPDOID) Then
                                Value = vbChecked
                            Else
                                Value = vbUnchecked
                            End If
                        Else
                            Value = vbUnchecked
                        End If
                    Else
                        Value = vbUnchecked
                    End If
                Else
                    Value = vbUnchecked
                End If
            Case GRD1_COL_TRAILER
                ' Trailer Number
                Value = .szTlrNr
               ' Trailer Message Indicator
                If .lTlrMsgId <> 0 Then
                   grdBayDetails(i).Columns(GRD1_COL_TRAILER).ForeColor = FDR_HIGHLIGHT_COLOR
                    grdBayDetails(i).Columns(GRD1_COL_TRAILER).Font.Bold = True
                Else
                    grdBayDetails(i).Columns(GRD1_COL_TRAILER).ForeColor = CLR_NEUTRAL
                    grdBayDetails(i).Columns(GRD1_COL_TRAILER).Font.Bold = False
                End If
           
            Case GRD1_COL_TYPE
                ' Trailer Type
                Value = .szEqpTyp
           
            Case GRD1_COL_PCS
                ' Pieces-Percent
                Value = .lPieces
               
                ' Set volume font underline
                If .szPcsPctFlg = QY_PCS Then
                    grdBayDetails(i).Columns(GRD1_COL_PCS).Font.Underline = True
                    grdBayDetails(i).Columns(GRD1_COL_PER).Font.Underline = False
                ElseIf .szPcsPctFlg = QY_PCT Then
                    grdBayDetails(i).Columns(GRD1_COL_PCS).Font.Underline = False
                    grdBayDetails(i).Columns(GRD1_COL_PER).Font.Underline = True
                Else
                    grdBayDetails(i).Columns(GRD1_COL_PCS).Font.Underline = False
                    grdBayDetails(i).Columns(GRD1_COL_PER).Font.Underline = False
                End If
            Case GRD1_COL_PER
                ' Pieces-Percent
                Value = .iPercent
           
            Case GRD1_COL_ORIG
                ' Load Name
                If .lLdMsgId <> 0 Then
                    grdBayDetails(i).Columns(GRD1_COL_ORIG).ForeColor = FDR_HIGHLIGHT_COLOR
                    grdBayDetails(i).Columns(GRD1_COL_OS).ForeColor = FDR_HIGHLIGHT_COLOR
                    grdBayDetails(i).Columns(GRD1_COL_DEST).ForeColor = FDR_HIGHLIGHT_COLOR
                    grdBayDetails(i).Columns(GRD1_COL_DS).ForeColor = FDR_HIGHLIGHT_COLOR
                    grdBayDetails(i).Columns(GRD1_COL_SEQ).ForeColor = FDR_HIGHLIGHT_COLOR
               
                    grdBayDetails(i).Columns(GRD1_COL_ORIG).Font.Bold = True
                    grdBayDetails(i).Columns(GRD1_COL_OS).Font.Bold = True
                    grdBayDetails(i).Columns(GRD1_COL_DEST).Font.Bold = True
                    grdBayDetails(i).Columns(GRD1_COL_DS).Font.Bold = True
                    grdBayDetails(i).Columns(GRD1_COL_SEQ).Font.Bold = True
                Else
                    grdBayDetails(i).Columns(GRD1_COL_ORIG).ForeColor = CLR_NEUTRAL
                    grdBayDetails(i).Columns(GRD1_COL_OS).ForeColor = CLR_NEUTRAL
                    grdBayDetails(i).Columns(GRD1_COL_DEST).ForeColor = CLR_NEUTRAL
                    grdBayDetails(i).Columns(GRD1_COL_DS).ForeColor = CLR_NEUTRAL
                    grdBayDetails(i).Columns(GRD1_COL_SEQ).ForeColor = CLR_NEUTRAL
               
                    grdBayDetails(i).Columns(GRD1_COL_ORIG).Font.Bold = False
                    grdBayDetails(i).Columns(GRD1_COL_OS).Font.Bold = False
                    grdBayDetails(i).Columns(GRD1_COL_DEST).Font.Bold = False
                    grdBayDetails(i).Columns(GRD1_COL_DS).Font.Bold = False
                    grdBayDetails(i).Columns(GRD1_COL_SEQ).Font.Bold = False
                End If
               
                ' Display Original load name (Sav) for LC 'E's already PD'd
                ' with another job
               
                ' Get Current PDData
                If .szDestinSrt = LD_CD_EMPTY And _
                        Not IsBlank(.szCurPDJobNa) And _
                        (pInst.PDInfo.szPDJobNr <> .szCurPDJobNa Or _
                          pInst.PDInfo.szPDJobDomSlic <> .szCurPDJobDom Or _
                          pInst.PDInfo.szPDJobDomCn <> .szCurPDJobDomCn) Then
                    ' Destination sort is Empty  and
                    ' job already PD'd and
                    ' PD'd with Different Job
                    If Len(Trim$(.szBayNr)) > 0 And Len(Trim$(.szSavOrigin)) > 0 Then
                        Value = MergeSabCny(.szSavOrigin, _
                                            .szSavOriginCn, _
                                            pInst.PDInfo.szCurrentCny)
                    Else
                        Value = MergeSabCny(.szOrigin, .szOriginCn, pInst.PDInfo.szCurrentCny)
                    End If
               
                Else ' Normal Display of load name
                    Value = MergeSabCny(.szOrigin, _
                                        .szOriginCn, _
                                        pInst.PDInfo.szCurrentCny)
                End If
            Case GRD1_COL_OS
                ' Get Current PDData
                If .szDestinSrt = LD_CD_EMPTY And _
                        Not IsBlank(.szCurPDJobNa) And _
                        (pInst.PDInfo.szPDJobNr <> .szCurPDJobNa Or _
                          pInst.PDInfo.szPDJobDomSlic <> .szCurPDJobDom Or _
                          pInst.PDInfo.szPDJobDomCn <> .szCurPDJobDomCn) Then
                    ' Destination sort is Empty  and
                    ' job already PD'd and
                    ' PD'd with Different Job
                    If Len(Trim$(.szBayNr)) > 0 And Len(Trim$(.szSavOriginSrt)) > 0 Then
                        Value = .szSavOriginSrt
                    Else
                        Value = .szOriginSrt
                    End If
                Else ' Normal Display of load name
                    Value = .szOriginSrt
                End If
            Case GRD1_COL_DEST
                ' Get Current PDData
                If .szDestinSrt = LD_CD_EMPTY And _
                        Not IsBlank(.szCurPDJobNa) And _
                        (pInst.PDInfo.szPDJobNr <> .szCurPDJobNa Or _
                          pInst.PDInfo.szPDJobDomSlic <> .szCurPDJobDom Or _
                          pInst.PDInfo.szPDJobDomCn <> .szCurPDJobDomCn) Then
                    ' Destination sort is Empty  and
                    ' job already PD'd and
                    ' PD'd with Different Job
                    If Len(Trim$(.szBayNr)) > 0 And Len(Trim$(.szSavDestin)) > 0 Then
                        Value = MergeSabCny(.szSavDestin, _
                                            .szSavDestinCn, _
                                            pInst.PDInfo.szCurrentCny)
                    Else
                        Value = MergeSabCny(.szDestin, .szDestinCn, pInst.PDInfo.szCurrentCny)
                    End If
   
                Else ' Normal Display of load name
                    Value = MergeSabCny(.szDestin, _
                                         .szDestinCn, _
                                         pInst.PDInfo.szCurrentCny)
                End If
            Case GRD1_COL_DS
                ' Get Current PDData
                If .szDestinSrt = LD_CD_EMPTY And _
                        Not IsBlank(.szCurPDJobNa) And _
                        (pInst.PDInfo.szPDJobNr <> .szCurPDJobNa Or _
                          pInst.PDInfo.szPDJobDomSlic <> .szCurPDJobDom Or _
                          pInst.PDInfo.szPDJobDomCn <> .szCurPDJobDomCn) Then
                    ' Destination sort is Empty  and
                    ' job already PD'd and
                    ' PD'd with Different Job
                    If Len(Trim$(.szBayNr)) > 0 And Len(Trim$(.szSavDestinSrt)) > 0 Then
                        Value = .szSavDestinSrt
                    Else
                        Value = .szDestinSrt
                    End If
               
                Else ' Normal Display of load name
                    Value = .szDestinSrt    '  QS 1/6/04
                End If
            Case GRD1_COL_SEQ
                If .szDestinSrt = LD_CD_EMPTY And _
                        Not IsBlank(.szCurPDJobNa) And _
                        (pInst.PDInfo.szPDJobNr <> .szCurPDJobNa Or _
                          pInst.PDInfo.szPDJobDomSlic <> .szCurPDJobDom Or _
                          pInst.PDInfo.szPDJobDomCn <> .szCurPDJobDomCn) Then
                    If Len(Trim$(.szBayNr)) > 0 And .iSavSequenceNr = 0 Then
                        If Len(CStr(Trim(.iSavSequenceNr))) = 1 Then
                           Value = "0" & .iSavSequenceNr
                        Else
                           Value = .iSavSequenceNr
                        End If
                    Else
                        If Len(CStr(Trim(.iSequenceNr))) = 1 Then
                           Value = "0" & .iSequenceNr
                        Else
                           Value = .iSequenceNr
                        End If
                    End If
                Else
                    If Len(CStr(Trim(.iSequenceNr))) = 1 Then
                       Value = "0" & .iSequenceNr
                    Else
                       Value = .iSequenceNr
                    End If
                End If
            Case GRD1_COL_LDCD
                If .szMultLdIr = IR_TRUE Then
               
                'Added by QS 3/7/05
                   Select Case sServiceTypeCD
                   Case NDASERVICE
                      grdBayDetails(i).Columns(GRD1_COL_LDCD).ForeColor = vbRed
                   Case SDASERVICE
                      grdBayDetails(i).Columns(GRD1_COL_LDCD).ForeColor = vbBlue
                   Case GRDSERVICE
                      grdBayDetails(i).Columns(GRD1_COL_LDCD).ForeColor = vbBrown
                   Case THREEDSSERVICE
                      grdBayDetails(i).Columns(GRD1_COL_LDCD).ForeColor = vbOrange
                   Case EAMSERVICE
                      grdBayDetails(i).Columns(GRD1_COL_LDCD).ForeColor = vbRed
                   End Select
                   grdBayDetails(i).Columns(GRD1_COL_LDCD).Font.Bold = True
                Else ' Set to default color
                    grdBayDetails(i).Columns(GRD1_COL_LDCD).ForeColor = CLR_NEUTRAL
                    grdBayDetails(i).Columns(GRD1_COL_LDCD).Font.Bold = False
                End If
                Value = .szLdCd
            Case GRD1_COL_RI
                ' Route Code
                'TT003628 - check to see if the load name is the same as a previous load name (Orig, Orig S, Dest, Dest S)
                If i = 1 Then
                    If (pInst.PDInfo.hPDHll(1).szOrigin = gpData.szOrigin And pInst.PDInfo.hPDHll(1).szOriginCn = gpData.szOriginCn And _
                        pInst.PDInfo.hPDHll(1).szDestin = gpData.szDestin And pInst.PDInfo.hPDHll(1).szDestinCn = gpData.szDestinCn And _
                        pInst.PDInfo.hPDHll(1).szDestinSrt = gpData.szDestinSrt And pInst.PDInfo.hPDHll(1).szOriginSrt = gpData.szOriginSrt) Then
                        .SameLoadName = True
                        If pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False Then .szRtId = pInst.PDInfo.hPDHll(1).szRtId
                        pInst.PDInfo.hPDHll(1).NextSameLoadName = True
                    End If
                ElseIf i = 2 Then
                    If (pInst.PDInfo.hPDHll(1).szOrigin = gpData.szOrigin And pInst.PDInfo.hPDHll(1).szOriginCn = gpData.szOriginCn And _
                        pInst.PDInfo.hPDHll(1).szDestin = gpData.szDestin And pInst.PDInfo.hPDHll(1).szDestinCn = gpData.szDestinCn And _
                        pInst.PDInfo.hPDHll(1).szDestinSrt = gpData.szDestinSrt And pInst.PDInfo.hPDHll(1).szOriginSrt = gpData.szOriginSrt) Then
                        .SameLoadName = True
                        pInst.PDInfo.hPDHll(1).NextSameLoadName = True
                        If pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False Then .szRtId = pInst.PDInfo.hPDHll(1).szRtId
                    End If
                   
                    If .SameLoadName = False Then
                        If (pInst.PDInfo.hPDHll(2).szOrigin = gpData.szOrigin And pInst.PDInfo.hPDHll(2).szOriginCn = gpData.szOriginCn And _
                            pInst.PDInfo.hPDHll(2).szDestin = gpData.szDestin And pInst.PDInfo.hPDHll(2).szDestinCn = gpData.szDestinCn And _
                            pInst.PDInfo.hPDHll(2).szDestinSrt = gpData.szDestinSrt And pInst.PDInfo.hPDHll(2).szOriginSrt = gpData.szOriginSrt) Then
                            .SameLoadName = True
                            pInst.PDInfo.hPDHll(2).NextSameLoadName = True
                            If pInst.PDInfo.hPDHll(2).InvalidLoadRouting = False Then .szRtId = pInst.PDInfo.hPDHll(2).szRtId
                        End If
                    End If
                End If
                Value = .szRtId
            Case GRD1_COL_CREATE
                ' Create Date
                Value = Left$(.szDspCrtDt, 2) & "/" & Right$(.szDspCrtDt, 2)
            Case GRD1_COL_DUE
                 ' Due Date
                If Not IsNull(.szDspDueDt) And .szDspDueDt <> "    " Then
                   Value = Left$(.szDspDueDt, 2) & "/" & Right$(.szDspDueDt, 2)
                Else
                   Value = " "
                End If
            Case GRD1_COL_DEP_TM
                ' Scheduled Departure Time
                If Not IsNull(.szDspScdDptTm) And .szDspScdDptTm <> "    " Then
                   Value = Left$(.szDspScdDptTm, 2) & ":" & Right$(.szDspScdDptTm, 2)
                Else
                   Value = ""
                End If
'            Case GRD1_COL_HAZMAT
'                '  Haz Mat
'                Value = .szHzMt
            Case GRD1_COL_TAG
                Value = .szTgCd
            Case GRD1_COL_REMARKS
                Value = .szRemarks
            End Select
        End With
    End If
   
     If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdBayDetails_ClassicRead Col = " & Col & "- End  - Index: " & Index)
    End If
   
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "grdBayDetails_ClassicRead"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                    IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                    Err.Description & " Module:" & gsError, _
                                    oProc, _
                                    ERROR_MSG, _
                                    FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
    End If
End Sub
 
Private Sub grdBayDetails_ColResize(Index As Integer, ByVal ColIndex As Integer, Cancel As Integer)
   If g_iDebug = 13 Then
           Call InfoLog("frmPDD grdBayDetails_ColResize - ColIndex: " & ColIndex & " - Begin - Index: " & Index)
   End If
   Cancel = True
   If g_iDebug = 13 Then
           Call InfoLog("frmPDD grdBayDetails_ColResize - ColIndex: " & ColIndex & " - End - Index: " & Index)
   End If
End Sub
 
Private Sub grdBayDetails_DblClick(Index As Integer)
If g_iDebug = 13 Then
     Call InfoLog("frmPDD grdBayDetails_DblClick Index: " & Index & " - Begin")
End If
If bPredispatchEmpty(Index) Then
   Select Case grdBayDetails(Index).Col
      Case GRD1_COL_TYPE
        grdBayDetails(Index).Columns(GRD1_COL_TYPE).DropDown = drpValTrailerType(Index)
    End Select
End If
If g_iDebug = 13 Then
    Call InfoLog("frmPDD grdBayDetails_DblClick Index: " & Index & " - End")
End If
End Sub
 
Private Sub grdBayDetails_GotFocus(Index As Integer)
If g_iDebug = 13 Then
    Call InfoLog("frmPDD grdBayDetails_GotFocus Index: " & Index & " - Begin")
End If
If bGridGotFocus = True Then
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdBayDetails_GotFocus Index: " & Index & " - End - If bGridGotFocus = True")
    End If
    Exit Sub
End If
Call SetGridForEdit(Index)
'bGridGotFocus = True
If (Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)) > 0 Then
    grdBayDetails(Index).SelStart = 0
    grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)
End If
        If g_iDebug = 13 Then
                Call InfoLog("frmPDD grdBayDetails_GotFocus Index: " & Index & " - End")
        End If
End Sub
 
Private Sub OLDgrdBayDetails_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   
  bDuplicateLoad = False
  Dim lRow As Long, lNewRow As Long
  Dim iCancel As Integer
  On Error GoTo Error_Handler
  If g_iDebug = 13 Then
      Call InfoLog("frmPDD OLDgrdBayDetails_KeyDown Index: " & Index & " - Begin")
  End If
  If bPredispatchEmpty(Index) = True Then
     If KeyCode = GRID_TAB Or KeyCode = GRID_RIGHT_ARROW Or KeyCode = GRID_LEFT_ARROW Or _
       KeyCode = GRID_DOWN_ARROW Or KeyCode = GRID_UP_ARROW Or KeyCode = GRID_ENTER Or _
       KeyCode = GRID_HOME Or KeyCode = GRID_END Or KeyCode = GRID_PGDN Or KeyCode = GRID_PGUP _
       Then
     
          If ctlWarningPopup.WarningVisible Then
            KeyCode = 0
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD OLDgrdBayDetails_KeyDown Index: " & Index & " - End - If ctlWarningPopup.WarningVisible")
            End If
            Exit Sub
          End If
      End If
   
      With grdBayDetails(Index)
        If (KeyCode = GRID_TAB And Shift <> 1) Or KeyCode = GRID_RIGHT_ARROW Or KeyCode = 13 Then
          Select Case .Col
            Case GRD1_COL_TYPE
                .Col = GRD1_COL_DEST
            Case GRD1_COL_DEST
                .Col = GRD1_COL_DS
            Case GRD1_COL_DS
                .Col = GRD1_COL_RI
            Case GRD1_COL_RI
                txtDolly(Index).Enabled = True
                txtDolly(Index).SetFocus
          End Select
          KeyCode = 0
        ElseIf (KeyCode = GRID_TAB) Or KeyCode = GRID_LEFT_ARROW Then
          Select Case .Col
            Case GRD1_COL_TYPE
                .Col = GRD1_COL_TYPE
            Case GRD1_COL_DEST
                .Col = GRD1_COL_TYPE
            Case GRD1_COL_DS
              .Col = GRD1_COL_DEST
            Case GRD1_COL_RI
              .Col = GRD1_COL_DS
          End Select
          KeyCode = 0
        End If
    End With
  End If
If g_iDebug = 13 Then
     Call InfoLog("frmPDD OLDgrdBayDetails_KeyDown Index: " & Index & " - End")
End If
Exit Sub
 
Error_Handler:
  gsError = "OLDgrdBayDetails_KeyDown"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
  End If
  Call InfoLog("frmPDD OLDgrdBayDetails_KeyDown Index: " & Index & " - Error Occured: " & Err.Description)
End Sub
 
Private Sub grdBayDetails_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
bDuplicateLoad = False
Dim lRow As Long, lNewRow As Long
Dim iCancel As Integer
Dim iCol As Integer
Dim iSplit As Integer
 
On Error GoTo Error_Handler
 
If g_iDebug = 13 Then
     Call InfoLog("frmPDD grdBayDetails_KeyDown Index: " & Index & " - Begin")
End If
If bPredispatchEmpty(Index) = True Then
     If KeyCode = GRID_TAB Or KeyCode = GRID_RIGHT_ARROW Or KeyCode = GRID_LEFT_ARROW Or _
       KeyCode = GRID_DOWN_ARROW Or KeyCode = GRID_UP_ARROW Or KeyCode = GRID_ENTER Or _
       KeyCode = GRID_HOME Or KeyCode = GRID_END Or KeyCode = GRID_PGDN Or KeyCode = GRID_PGUP _
       Or KeyCode = GRID_ENTER Then
        If grdBayDetails(Index).Col = GRD1_COL_DEST Or _
           grdBayDetails(Index).Col = GRD1_COL_DS Or _
           grdBayDetails(Index).Col = GRD1_COL_TYPE Then
            iSplit = grdBayDetails(Index).Split
            iCol = grdBayDetails(Index).Col
            Call grdBayDetails_BeforeRowColChange(Index, iCancel)
            If CBool(iCancel) = True Then
                grdBayDetails(Index).Split = iSplit
                grdBayDetails(Index).Col = iCol
            End If
        End If
        If ctlWarningPopup.WarningVisible Then
            KeyCode = 0
            grdBayDetails(Index).Enabled = True
            grdBayDetails(Index).SetFocus
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD grdBayDetails_KeyDown Index: " & Index & " - End - If ctlWarningPopup.WarningVisible")
            End If
            Exit Sub
        End If
      ' TT000394 - ESC key is not working
      'ElseIf KeyCode = oKeyDefs.Escape_Key Then
      '  Call ClearField
     End If
     With grdBayDetails(Index)
       If (KeyCode = GRID_TAB And Shift <> 1) Or KeyCode = GRID_RIGHT_ARROW Or KeyCode = 13 Then
         Select Case .Col
            Case GRD1_COL_TYPE
            If m_bDropDownTlrTyp(Index) Then
                If Not IsNull(drpValTrailerType(Index).SelectedItem) Then
              'If item in the drop down is selected and <enter> is pressed, update grid
                    .Columns.Item(GRD1_COL_TYPE).Value = drpValTrailerType(Index).Text
                End If
            End If
        End Select
       ElseIf KeyCode = GRID_DOWN_ARROW Then
             If .Col = GRD1_COL_TYPE Then
                drpValTrailerType(Index).Enabled = True
                drpValTrailerType(Index).SetFocus
                m_bDropDownTlrTyp(Index) = True
             End If
       End If
     End With
  End If
  If g_iDebug = 13 Then
     Call InfoLog("frmPDD grdBayDetails_KeyDown Index: " & Index & " - End")
  End If
Exit Sub
 
Error_Handler:
  gsError = "grdBayDetails_KeyDown"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
  End If
  Call InfoLog("frmPDD grdBayDetails_KeyDown Index:" & Index & " - Error Occured: " & Err.Description)
End Sub
 
Private Sub grdBayDetails_LostFocus(Index As Integer)
100 On Error GoTo Error_Handler
101 If g_iDebug = 13 Then
102    Call InfoLog("frmPDD grdBayDetails_LostFocus Index:" & Index & " - Begin")
103 End If
104 If bGridGotFocus = True Then
        grdBayDetails(Index).Enabled = True
105     grdBayDetails(Index).SetFocus
106     If g_iDebug = 13 Then
107         Call InfoLog("frmPDD grdBayDetails_LostFocus Index: " & Index & " - End - If bGridGotFocus = True")
108     End If
109     Exit Sub
110 End If
 
111 If g_bRoutingOpen = True Or (bTrailerError = True And bEscPressedOnError = True) Then
  '  iBayIndex = Index
112     If g_iDebug = 13 Then
113         Call InfoLog("frmPDD grdBayDetails_LostFocus Index: " & Index & " - End - If g_bRoutingOpen = True Or...")
114     End If
115     Exit Sub
116 End If
   
117 If bPredispatchEmpty(Index) Then
        '09/22/14 - do not want to set focus to dolly when saving
118     If (ctlWarningPopup.Visible = False Or bLoadRoutingError) And bSaveData = False Then
119            txtDolly(Index).Visible = True
120            txtDolly(Index).Enabled = True
               If g_iDebug = 13 Then
                    Call InfoLog("frmPDD grdBayDetails_LostFocus Enabled = " & txtDolly(Index).Enabled & "ctlWarningPopup.Visible = " & ctlWarningPopup.Visible & _
                                 "bLoadRouting = " & bLoadRoutingError & " bSaveData = " & bSaveData & " txtDolly Visible = " & txtDolly(Index).Visible & " Index = " & Index)
               End If
121            txtDolly(Index).SetFocus
        End If
       
151     If (grdBayDetails(Index).Col = GRD1_COL_DEST Or grdBayDetails(Index).Col = GRD1_COL_DS Or grdBayDetails(Index).Col = GRD1_COL_TYPE) And Not IsUserShiftTabbing And mbClearData = False Then
152         iBayIndex = Index
153     End If
122     UpdatePredispatchEmptyData (Index)
 
123 ElseIf bPredispatchEmptyWithTrailer(Index) Then
124     If (grdBayDetails(Index).Col = GRD1_COL_DEST Or grdBayDetails(Index).Col = GRD1_COL_DS) And Not IsUserShiftTabbing And mbClearData = False Then
125         If grdBayDetails(Index).Enabled = False Then grdBayDetails(Index).Enabled = True
126         grdBayDetails(Index).SetFocus
130         iBayIndex = Index
131         If (Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)) > 0 Then
132             grdBayDetails(Index).SelStart = 0
133             grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)
134         End If
135         If g_iDebug = 13 Then
136             Call InfoLog("frmPDD grdBayDetails_LostFocus Index:" & Index & " - End - ElseIf bPredispatchEmptyWithTrailer")
137         End If
138         Exit Sub
139     ElseIf mbClearData = False And ctlWarningPopup.WarningVisible = False And bSaveData = False Then
140         txtDolly(Index).Enabled = True
145         txtDolly(Index).SetFocus
146     ElseIf mbClearData = True And ctlWarningPopup.WarningVisible = False And bSaveData = False Then
147         ClearField
148     ElseIf ctlWarningPopup.WarningVisible = True And bSaveData = False Then
149         If grdBayDetails(Index).Col = GRD1_COL_RI Then
150             If grdBayDetails(Index).Enabled = False Then grdBayDetails(Index).Enabled = True
156             grdBayDetails(Index).SetFocus
157             iBayIndex = Index
158             If (Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)) > 0 Then
159                 grdBayDetails(Index).SelStart = 0
160                 grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)
161             End If
162         End If
163     End If
164 End If
 
165 Call SetGridNormal(Index)
166 If g_iDebug = 13 Then
167      Call InfoLog("frmPDD grdBayDetails_LostFocus Index: " & Index & " - End")
168 End If
Exit Sub
Error_Handler:
  gsError = "grdBayDetails_LostFocus"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError & " Line: " & Erl, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
  End If
End Sub
 
Private Sub grdBayDetails_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim bInvalidLoadRouting As Boolean
If g_iDebug = 13 Then
     Call InfoLog("frmPDD grdBayDetails_RowColChange Index:" & Index & " - Begin")
End If
 
On Error GoTo Error_Handler
bInvalidLoadRouting = False
If LastCol > -1 Then
grdBayDetails(Index).Columns(LastCol).DropDown = vbNullString
End If
 
'TT003628
If pInst.PDInfo.hPDHll.Count > Index And Index > 0 Then
    If pInst.PDInfo.hPDHll(Index + 1).SameLoadName = True And grdBayDetails(Index).Col = GRD1_COL_RI Then
        If Index = 1 Then
            If (pInst.PDInfo.hPDHll(1).szOrigin = pInst.PDInfo.hPDHll(Index + 1).szOrigin And pInst.PDInfo.hPDHll(1).szOriginCn = pInst.PDInfo.hPDHll(Index + 1).szOriginCn And _
                pInst.PDInfo.hPDHll(1).szDestin = pInst.PDInfo.hPDHll(Index + 1).szDestin And pInst.PDInfo.hPDHll(1).szDestinCn = pInst.PDInfo.hPDHll(Index + 1).szDestinCn And _
                pInst.PDInfo.hPDHll(1).szDestinSrt = pInst.PDInfo.hPDHll(Index + 1).szDestinSrt And pInst.PDInfo.hPDHll(1).szOriginSrt = pInst.PDInfo.hPDHll(Index + 1).szOriginSrt And _
                pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False) Then
                pInst.PDInfo.hPDHll(Index + 1).szRtId = pInst.PDInfo.hPDHll(1).szRtId
                grdBayDetails(Index).Columns(GRD1_COL_RI).Text = pInst.PDInfo.hPDHll(Index + 1).szRtId
                pInst.PDInfo.hPDHll(Index + 1).InvalidLoadRouting = pInst.PDInfo.hPDHll(1).InvalidLoadRouting
                pInst.PDInfo.hPDHll(Index + 1).SendPredispatchPreforecast = pInst.PDInfo.hPDHll(1).SendPredispatchPreforecast
                pInst.PDInfo.hPDHll(Index + 1).EmptyRouteCodeChanged = pInst.PDInfo.hPDHll(1).EmptyRouteCodeChanged
            ElseIf pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True Then
                bInvalidLoadRouting = True
            End If
        ElseIf Index = 2 Then
            If (pInst.PDInfo.hPDHll(1).szOrigin = pInst.PDInfo.hPDHll(Index + 1).szOrigin And pInst.PDInfo.hPDHll(1).szOriginCn = pInst.PDInfo.hPDHll(Index + 1).szOriginCn And _
                pInst.PDInfo.hPDHll(1).szDestin = pInst.PDInfo.hPDHll(Index + 1).szDestin And pInst.PDInfo.hPDHll(1).szDestinCn = pInst.PDInfo.hPDHll(Index + 1).szDestinCn And _
                pInst.PDInfo.hPDHll(1).szDestinSrt = pInst.PDInfo.hPDHll(Index + 1).szDestinSrt And pInst.PDInfo.hPDHll(1).szOriginSrt = pInst.PDInfo.hPDHll(Index + 1).szOriginSrt And _
                pInst.PDInfo.hPDHll(1).InvalidLoadRouting = False) Then
                pInst.PDInfo.hPDHll(Index + 1).szRtId = pInst.PDInfo.hPDHll(1).szRtId
                grdBayDetails(Index).Columns(GRD1_COL_RI).Text = pInst.PDInfo.hPDHll(Index + 1).szRtId
                pInst.PDInfo.hPDHll(Index + 1).InvalidLoadRouting = pInst.PDInfo.hPDHll(1).InvalidLoadRouting
                pInst.PDInfo.hPDHll(Index + 1).SendPredispatchPreforecast = pInst.PDInfo.hPDHll(1).SendPredispatchPreforecast
                pInst.PDInfo.hPDHll(Index + 1).EmptyRouteCodeChanged = pInst.PDInfo.hPDHll(1).EmptyRouteCodeChanged
            ElseIf pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True Then
                bInvalidLoadRouting = True
            End If
           
            If (pInst.PDInfo.hPDHll(2).szOrigin = pInst.PDInfo.hPDHll(Index + 1).szOrigin And pInst.PDInfo.hPDHll(2).szOriginCn = pInst.PDInfo.hPDHll(Index + 1).szOriginCn And _
                pInst.PDInfo.hPDHll(2).szDestin = pInst.PDInfo.hPDHll(Index + 1).szDestin And pInst.PDInfo.hPDHll(2).szDestinCn = pInst.PDInfo.hPDHll(Index + 1).szDestinCn And _
                pInst.PDInfo.hPDHll(2).szDestinSrt = pInst.PDInfo.hPDHll(Index + 1).szDestinSrt And pInst.PDInfo.hPDHll(2).szOriginSrt = pInst.PDInfo.hPDHll(Index + 1).szOriginSrt And _
                pInst.PDInfo.hPDHll(2).InvalidLoadRouting = False) Then
                pInst.PDInfo.hPDHll(Index + 1).szRtId = pInst.PDInfo.hPDHll(2).szRtId
                grdBayDetails(Index).Columns(GRD1_COL_RI).Text = pInst.PDInfo.hPDHll(Index + 1).szRtId
                pInst.PDInfo.hPDHll(Index + 1).InvalidLoadRouting = pInst.PDInfo.hPDHll(2).InvalidLoadRouting
                pInst.PDInfo.hPDHll(Index + 1).SendPredispatchPreforecast = pInst.PDInfo.hPDHll(2).SendPredispatchPreforecast
                pInst.PDInfo.hPDHll(Index + 1).EmptyRouteCodeChanged = pInst.PDInfo.hPDHll(2).EmptyRouteCodeChanged
            ElseIf pInst.PDInfo.hPDHll(2).InvalidLoadRouting = True Then
                bInvalidLoadRouting = True
            End If
        End If
       
        If bInvalidLoadRouting = False And _
           Not (Me.ActiveControl.Name = "grdMultLegView" Or Me.ActiveControl.Name = "txtDriverName" Or Me.ActiveControl.Name = "txtJobNr") Then
          '  bGridGotFocus = False
            bBayCompeleted(Index) = True
            bBalNrValidated(Index) = True
            If bSaveData = False Then
                txtDolly(Index).Enabled = True
                txtDolly(Index).SetFocus
            End If
'            DoEvents
        End If
     '   bGridGotFocus = False
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD grdBayDetails_RowColChange Index:" & Index & " - End  -If pInst.PDInfo.hPDHll.Count > Index And Index > 0")
        End If
        Exit Sub
    End If
End If
 If bPredispatchEmpty(Index) Or bPredispatchEmptyWithTrailer(Index) Then
    Select Case grdBayDetails(Index).Col
        Case GRD1_COL_TYPE
            If bPredispatchEmpty(Index) Then
                grdBayDetails(Index).EditActive = True
                grdBayDetails(Index).SelStart = 0
                grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Splits(1).Columns(GRD1_COL_TYPE).Value)
                grdBayDetails(Index).Columns(GRD1_COL_TYPE).DropDown = drpValTrailerType(Index)
                grdBayDetails(Index).Columns(GRD1_COL_TYPE).AutoDropDown = True
           End If
        Case GRD1_COL_DEST
            grdBayDetails(Index).EditActive = True
            grdBayDetails(Index).SelStart = 0
            grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Splits(2).Columns(GRD1_COL_DEST).Value)
        Case GRD1_COL_DS
            grdBayDetails(Index).EditActive = True
            grdBayDetails(Index).SelStart = 0
            grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Splits(2).Columns(GRD1_COL_DS).Value)
        Case GRD1_COL_RI
            grdBayDetails(Index).EditActive = True
            grdBayDetails(Index).SelStart = 0
            grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Splits(3).Columns(GRD1_COL_RI).Value)
    End Select
   
    If grdBayDetails(Index).Col < 0 Then
        If LastCol = GRD1_COL_RI Then
            txtDolly(Index).Enabled = True
            txtDolly(Index).SetFocus
        End If
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD grdBayDetails_RowColChange Index:" & Index & " - End  - If bPredispatchEmpty(Index) Or...")
        End If
        Exit Sub
    End If
   
    If (Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)) > 0 Then
          grdBayDetails(Index).SelStart = 0
          grdBayDetails(Index).SelLength = Len(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Text)
    End If
ElseIf LastCol = GRD1_COL_RI Then
    If pInst.PDInfo.hPDHll.Count > 0 Then
        pInst.PDInfo.hPDHll(Index + 1).szRtId = grdBayDetails(Index).Columns(GRD1_COL_RI).Value
    End If
End If
If g_iDebug = 13 Then
    Call InfoLog("frmPDD grdBayDetails_RowColChange Index:" & Index & " - End")
End If
Exit Sub
 
Error_Handler:
  gsError = "grdBayDetails_RowColChange"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
End If
End Sub
 
Private Sub grdBayDetails_UnboundGetRelativeBookmark(Index As Integer, _
                                                     StartLocation As Variant, _
                                                     ByVal Offset As Long, _
                                                     NewLocation As Variant, _
                                                     ApproximatePosition As Long)
 
    On Error GoTo ErrorExit
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdBayDetails_UnboundGetRelativeBookmark - Begin: Index = " & Index)
    End If
 
' TDBGrid1 calls this routine each time it needs to
' reposition itself. StartLocation is a bookmark
' supplied by the grid to indicate the "current"
' position -- the row we are moving from. Offset is
' the number of rows we must move from StartLocation
' in order to arrive at the desired destination row.
' A positive offset means the desired record is after
' the StartLocation, and a negative offset means the
' desired record is before StartLocation.
' If StartLocation is NULL, then we are positioning
' from either BOF or EOF. Once we determine the
' correct index for StartLocation, we will simply add
' the offset to get the correct destination row.
' GetRelativeBookmark already does all of this, so we
' just call it here.
    NewLocation = GetRelativeBookmark(StartLocation, Offset, 1)
   
' If we are on a valid data row (i.e., not at BOF or
' EOF), then set the ApproximatePosition (the ordinal
' row number) to improve scroll bar accuracy. We can
' call IndexFromBookmark to do this.
    If Not IsNull(NewLocation) Then
       ApproximatePosition = IndexFromBookmark(NewLocation, 0, 1)
    End If
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdBayDetails_UnboundGetRelativeBookmark - End: Index = " & Index)
    End If
   
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "grdBayDetails_UnboundGetRelativeBookmark"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                    IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                    Err.Description & " Module:" & gsError, _
                                    oProc, _
                                    ERROR_MSG, _
                                    FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
    End If
End Sub
 
Private Sub SetELOCInfo()
    Dim iRc As Integer
    Dim bResult As Boolean
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetELOCInfo() - Begin")
    End If
   
    On Error GoTo Error_Handler
       txtEloc.Text = pInst.PDInfo.szPDEloc
           ' Find schedule information for JOB NUMBER
       If Not IsNull(pInst.PDInfo.SegmentSystemNumberOID) Then
        If (Len(Trim$(pInst.PDInfo.SegmentSystemNumberOID)) > 0) Then
           iRc = pInst.PDInfo.GetSchedTrailer()
   
           If iRc = DB_SUCCESS Then  ' Search the Returned Data
              iRc = pInst.SearchSchedInfo()
        '        If lResult = PD_SUCCESS Or lResult = PD_FAILURE Or _
        '            lResult = PD_EXIT Then
        '            GetSchedJobs = lResult
        '        Else
        '            ' If this is the case we want it treated as if GetSchedTrailer
        '            ' returned no data
        '            lDbReturnMsg = DB_EOF
        '        End If
           Else
               bResult = False
           End If
          
           If bLoadAssignedToLeg = False Then
               mbElocValid = pInst.RetrieveSchdMovement(pInst.PDInfo.szPDEloc, pInst.PDInfo.szPDElocCn)
           Else
               If pInst.PDInfo.hPDHll.Count > 0 And mbElocValid = True Then
                  Call DisplayLdData(pInst.PDInfo.hPDHll)
                  Dim i As Integer
                  For i = 0 To pInst.PDInfo.hPDHll.Count - 1
                    bDataFromEloc(i) = True
                  Next i
               End If
           End If
           If mbElocValid = False Then
               If pInst.PDInfo.hPDHll.Count > 0 Then
                   mbElocValid = True
                   Call DisplayLdData(pInst.PDInfo.hPDHll)
               End If
           End If
   
           bResult = mbElocValid
           If iMyRc = UMBID_DELETE Or iMyRc = DB_EOF Then   'Added for the condition of removing the load.
              SetAvailMenuAndTool MAT_BASE_DATA0, False, False
           Else  'for no load found  'QS 1/2/05
              If pInst.PDInfo.hPDHll.Count > 0 Then   '5/24
   
                   If pInst.PDInfo.hPDHll.Item(1).szMultLdIr = IR_TRUE And (pInst.PDInfo.hPDHll.Item(1).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(1).lLdMsgId > 0) Then
                      SetAvailMenuAndTool MAT_BASE_DATA0, True, True
                   ElseIf pInst.PDInfo.hPDHll.Item(1).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(1).lLdMsgId > 0 Then
                      SetAvailMenuAndTool MAT_BASE_DATA0, True, False
                   ElseIf pInst.PDInfo.hPDHll.Item(1).szMultLdIr = IR_TRUE Then
                      SetAvailMenuAndTool MAT_BASE_DATA0, False, True
                   Else
                      SetAvailMenuAndTool MAT_BASE_DATA0, False, False
                   End If
               Else
                  SetAvailMenuAndTool MAT_BASE_DATA0, False, False
               End If
               'SetAvailMenuAndTool MAT_BASE_DATA0, False, False
   
           End If
           Else
                MsgBox "An Error occurred. All information was not received for the leg. Try getting out of the predispatch screen and going back in to resolve, or try resending the leg information from SADE.", vbCritical
                Err.Raise 400, TypeName(Me), "An Error occurred. All information was not received for the leg. Try getting out of the predispatch screen and going back in to resolve, or try resending the leg information from SADE."
           End If
       Else
            MsgBox "An Error occurred. All information was not received for the leg. Try getting out of the predispatch screen and going back in to resolve, or try resending the leg information from SADE.", vbCritical
            Err.Raise 400, TypeName(Me), "An Error occurred. All information was not received for the leg. Try getting out of the predispatch screen and going back in to resolve, or try resending the leg information from SADE."
       End If
       If g_iDebug = 13 Then
            Call InfoLog("frmPDD SetELOCInfo() - End")
       End If
    Exit Sub
      
Error_Handler:
  glErrNum = 400
  gsError = "SetELOCInfo()"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER, NO_POP_UP) Then
'        Set oEventlog = Nothing
'        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'        Unload Me
        Err.Raise Err.Number, Err.Source, Err.Description
End If
End Sub
 
Private Sub grdMultLegView_BeforeRowColChange(Cancel As Integer)
    On Error GoTo ErrorHandler
   
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_BeforeRowColChange - Begin")
    End If
    If IsNull(grdMultLegView.DestinationRow) Or _
       grdMultLegView.DestinationRow < 0 Or _
       grdMultLegView.DestinationRow = "" Then Exit Sub
   
    If xaMultiLegArray.Value(grdMultLegView.DestinationRow, 8) = True Then
        Cancel = True
    End If
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_BeforeRowColChange - End")
    End If
    Exit Sub
      
ErrorHandler:
  glErrNum = 400
  gsError = "grdMultLegView_GotFocus()"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
     '   Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
  End If
End Sub
 
Private Sub grdMultLegView_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_FetchRowStyle - Begin")
    End If
    If Not IsNull(Bookmark) And xaMultiLegArray.UpperBound(1) <> -1 Then
        If xaMultiLegArray.Value(Bookmark, 8) = True Then
            RowStyle.Locked = True
            RowStyle.ForeColor = vbGrayText
 
            If Not (grdMultLegView.Bookmark < 0 Or IsNull(grdMultLegView.Bookmark)) Then
               If xaMultiLegArray.Value(grdMultLegView.Bookmark, 8) = True Then
            '   If xaMultiLegArray.UpperBound(1) = 0 Or (Bookmark = xaMultiLegArray.UpperBound(1)) Then
                   grdMultLegView.MarqueeStyle = dbgNoMarquee
                '   txtJobNr.SetFocus
               Else
                   grdMultLegView.MarqueeStyle = dbgHighlightRow
               End If
            End If
        Else
            grdMultLegView.MarqueeStyle = dbgHighlightRow
        End If
    End If
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_FetchRowStyle - End")
    End If
End Sub
 
Private Sub grdMultLegView_GotFocus()
    On Error GoTo ErrorHandler
   
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_GotFocus() - Begin")
    End If
    TextEnd.Enabled = True
    TextEnd.TabStop = True
    grdMultLegView.Styles(5).BackColor = vbGreen
    If pInst.PDInfo.hPDHll.Count > 0 And g_bShowMatchingLoad = True Then
       Call DeleteLoadOnScreen(0, True)
    End If
    If g_bShowMatchingLoad = True Then
        SetELOCInfo
    End If
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_GotFocus() - End")
    End If
    Exit Sub
      
ErrorHandler:
  glErrNum = 400
  gsError = "grdMultLegView_GotFocus()"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER, NO_POP_UP) Then
      '  Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
       ClosefrmPDD
  End If
End Sub
 
Private Sub grdMultLegView_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_KeyDown - Begin")
    End If
   
    On Error GoTo Error_Handler
   
    Select Case KeyCode
        Case oKeyDefs.Enter_Key
            EnterAsTab (KeyCode)
        Case vbKeyDown
            If grdMultLegView.Bookmark + 1 <> xaMultiLegArray.Count(1) Then
                If xaMultiLegArray.Value(grdMultLegView.Bookmark + 1, 8) = True Then
                    If grdMultLegView.Bookmark + 1 < xaMultiLegArray.Count(1) Then
                        i = grdMultLegView.Bookmark + 1
                        Do While i <> xaMultiLegArray.Count(1)
                            If xaMultiLegArray.Value(i, 8) = True Then
                                i = i + 1
                            Else
                                Exit Do
                            End If
                        Loop
                        If i <> xaMultiLegArray.Count(1) Then
                            grdMultLegView.Bookmark = i
                            KeyCode = 0
                        Else
                            KeyCode = 0
                        End If
                    Else
                        KeyCode = 0
                    End If
                End If
            End If
        Case vbKeyUp
            If grdMultLegView.Bookmark <> 0 Then
                If xaMultiLegArray.Value(grdMultLegView.Bookmark - 1, 8) = True Then
                    If grdMultLegView.Bookmark - 1 > 0 Then
                        i = grdMultLegView.Bookmark - 1
                        Do While i <> 0
                            If xaMultiLegArray(i, 8) = True Then
                                i = i - 1
                            Else
                                Exit Do
                            End If
                        Loop
                        If i <> 0 Then
                            grdMultLegView.Bookmark = i
                        Else
                            KeyCode = 0
                        End If
                    Else
                        KeyCode = 0
                    End If
                End If
            End If
        Case vbKeyHome
            If grdMultLegView.Bookmark <> 0 Then
                If xaMultiLegArray.Value(0, 8) = False Then
                    grdMultLegView.MoveFirst
                Else
                   For i = 1 To xaMultiLegArray.UpperBound(1)
                        If xaMultiLegArray.Value(i, 8) = False Then
                            grdMultLegView.Bookmark = i
                            Exit For
                        End If
                   Next i
                  
                   If i = xaMultiLegArray.UpperBound(1) Then
                        KeyCode = 0
                   End If
                  
                End If
            End If
        Case vbKeyEnd
            If xaMultiLegArray.Value(xaMultiLegArray.UpperBound(1) - 1, 8) = False Then
                grdMultLegView.MoveLast
            Else
                For i = xaMultiLegArray.UpperBound(1) - 1 To 0 Step -1
                    If xaMultiLegArray.Value(i, 8) = False Then
                        grdMultLegView.Bookmark = i
                        Exit For
                    End If
                Next i
               
                If i < 0 Then
                    KeyCode = 0
                End If
            End If
    End Select
    If g_iDebug = 13 Then
         Call InfoLog("frmPDD grdMultLegView_KeyDown - End")
    End If
    Exit Sub
Error_Handler:
  glErrNum = 400
  gsError = "grdMultLegView_KeyDown"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                  Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
    '    Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Sub
 
Private Sub grdMultLegView_LostFocus()
10    If g_iDebug = 13 Then
11         Call InfoLog("frmPDD grdMultLegView_LostFocus() - Begin")
12    End If
13    On Error GoTo Error_Handler
14    If Not IsUserShiftTabbing Then
15        If grdMultLegView.Bookmark < 0 Or IsNull(grdMultLegView.Bookmark) Then
16            If g_iDebug = 13 Then
17                Call InfoLog("frmPDD grdMultLegView_LostFocus() - End - If grdMultLegView.Bookmark < 0 Or...")
18                Exit Sub
19            End If
20        End If
21        If xaMultiLegArray.UpperBound(1) <> -1 Then
 
31            If xaMultiLegArray.Value(0, 8) = True Then
32                If xaMultiLegArray.UpperBound(1) = 0 Then
        '            grdMultLegView.Bookmark = grdMultLegView.Bookmark + 1
        '        Else
                    'TT#3330 - app2dmw - 3-25-2013 - start
33                    frmBayData.Enabled = False
34                    txtLastDriverName = ""
35                    jobLegsUsedUp = True
                    'MsgBox "All legs for this job has been departed for this date. "
36                    MsgBox LoadResString(CLng(1113))
                    'TT#3330 - app2dmw - 3-25-2013 - end
37                    If grdMultLegView.Enabled = False Then grdMultLegView.Enabled = True
38                    grdMultLegView.SetFocus
39                    If g_iDebug = 13 Then
40                        Call InfoLog("frmPDD grdMultLegView_LostFocus() - End - If xaMultiLegArray.Value(0, 8) = True")
50                    End If
51                    Exit Sub
                'TT#3330 - app2dmw - 3-25-2013 - start
52                Else
53                    jobLegsUsedUp = False
                'TT#3330 - app2dmw - 3-25-2013 - end
54                End If
55            End If
56            Call txtEloc_LostFocus
          End If
57    End If
58    grdMultLegView.Styles(5).BackColor = vbYellow
59    If g_iDebug = 13 Then
60        Call InfoLog("frmPDD grdMultLegView_LostFocus() - End")
61    End If
62    Exit Sub
Error_Handler:
  glErrNum = 400
  gsError = "grdMultLegView_LostFocus"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
            IIf(Err.Number <> 0, Err.Number, glErrNum), _
            Err.Description & " Module:" & gsError & " Line:" & Erl, _
            oProc, _
            ERROR_MSG, _
            FEEDER_DISPATCH_DRIVER) Then
     MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
   '  Unload Me
  End If
End Sub
 
Private Sub grdMultLegView_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Error_Handler
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD grdMultLegView_RowColChange - Begin")
    End If
    If grdMultLegView.Bookmark < 0 Or IsNull(grdMultLegView.Bookmark) Or txtJobNr.Text = "" Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD grdMultLegView_RowColChange - End - If grdMultLegView.Bookmark < 0 Or...")
        End If
        Exit Sub
    End If
   
    If xaMultiLegArray.UpperBound(1) <> -1 Then
        'make sure that the current leg has been departed, and the last leg has been departed.  If the current leg has not been
        'departed, allow the current leg to be predispatched
        If xaMultiLegArray.Value(xaMultiLegArray.UpperBound(1), 8) = True And xaMultiLegArray(grdMultLegView.Bookmark, 8) = True Then
            'TT#3330 - app2dmw - 3-25-2013 - start
            frmBayData.Enabled = False
            txtLastDriverName = ""
            jobLegsUsedUp = True
            'MsgBox "All legs for this job has been departed for this date. "
            MsgBox LoadResString(CLng(1113))
            'TT#3330 - app2dmw - 3-25-2013 - end
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD grdMultLegView_RowColChange - End - If xaMultiLegArray.UpperBound(1) <> -1")
            End If
            Exit Sub
        'TT#3330 - app2dmw - 3-25-2013 - start
        Else
            jobLegsUsedUp = False
        'TT#3330 - app2dmw - 3-25-2013 - end
        End If
    Else
        If g_iDebug = 13 Then
                Call InfoLog("frmPDD grdMultLegView_RowColChange - End - If NOT xaMultiLegArray.UpperBound(1) <> -1")
        End If
        Exit Sub
    End If
   
    If Not pInst.bArriveFlag Then g_bShowMatchingLoad = True
   
    If pInst.PDInfo.hPDHll.Count > 0 Then
       Call DeleteLoadOnScreen(0, True)
    End If
   
    Call function_key_pressed(oKeyDefs.Enter_Key, 0)
    If g_bShowMatchingLoad Then SetELOCInfo
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdMultLegView_RowColChange - End")
    End If
    Exit Sub
Error_Handler:
glErrNum = 400
gsError = "grdMultLegView_RowColChange"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                           IIf(Err.Number <> 0, Err.Number, glErrNum), _
                           Err.Description & " Module:" & gsError, _
                           oProc, _
                           ERROR_MSG, _
                           FEEDER_DISPATCH_DRIVER, NO_POP_UP) Then
    MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    'Unload Me
    'grdMultLegView.Enabled = False
'     grdBayDetails(0).TabStop = False
'     grdBayDetails(1).TabStop = False
'     grdBayDetails(2).TabStop = False
'    ClosefrmPDD
 
End If
End Sub
 
Private Sub grdMultLegView_Scroll(Cancel As Integer)
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdMultLegView_Scroll - Begin")
    End If
    grdMultLegView.Refresh
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD grdMultLegView_Scroll - End")
    End If
End Sub
 
Private Sub mnuClose_Click()
'    Unload Me
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuClose_Click() - Begin")
    End If
    ClosefrmPDD '<WR01024><ADID:dmx7plm>
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuClose_Click() - End")
    End If
End Sub
 
Private Sub mnuDolliesDisplay_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuDolliesDisplay_Click() - Begin")
    End If
   GetDollyScreen txtEloc.Text
   If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuDolliesDisplay_Click() - End")
    End If
End Sub
 
Private Sub mnuF10Accept_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF10Accept_Click() - Begin")
    End If
   DoSaveData
   If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF10Accept_Click() - End")
    End If
End Sub
 
Private Sub mnuF11CancelPredispatch_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF11CancelPredispatch_Click() - Begin")
    End If
    SendCancelPredispatch
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF11CancelPredispatch_Click() - End")
    End If
End Sub
 
Private Sub mnuF2TrailerMessage_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF2TrailerMessage_Click() - Begin")
    End If
    ShowLdTlrMsg
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF2TrailerMessage_Click() - End")
    End If
End Sub
 
Private Sub mnuF3JobSchedules_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF3JobSchedules_Click() - Begin")
    End If
    DoDspSchedByDriver
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF3JobSchedules_Click() - End")
    End If
End Sub
 
Private Sub mnuF4MultiBayDisplay_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF4MultiBayDisplay_Click() - Begin")
    End If
    DoMultiBayDisplay
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF4MultiBayDisplay_Click() - End")
    End If
End Sub
 
Private Sub mnuF5QuickSort_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF5QuickSort_Click() - Begin")
    End If
    DoQuickSort
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF5QuickSort_Click() - End")
    End If
End Sub
 
Private Sub mnuF6AddUpdateLeg_Click()
    Dim oIVISEditor As IVISEditorStarter.IVISEditorDll
    Dim oWSInfo As HFCSWSInformation.clsWSInfo
    Dim sUserName As String
    Dim sIVISInfo As String
    Dim sDriverDomicile As String
   
    On Error GoTo ErrorHandler
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF6AddUpdateLeg_Click() - Begin")
    End If
   
    If (ctlToolbarManager.isButtonEnabled(BT_F6_UPD_LEG)) Then
        Set oIVISEditor = New IVISEditorStarter.IVISEditorDll
        Set oWSInfo = New HFCSWSInformation.clsWSInfo
       
        sUserName = Trim$(oWSInfo.HFCSUserName)
        sIVISInfo = oClsDb.GetSLICInfo
        'get region district information
        sDriverDomicile = oClsDb.GetDomicileSLICByNrAndRegionDistrict(pInst.PDInfo.szPDJobDomSlic, pInst.PDInfo.szPDJobDomCn)   ' pInst.PDInfo.szPDJobDomCn & "," & "," & "," & pInst.PDInfo.szPDJobDomSlic
       
        'If xaMultiLegArray.Value(xaMultiLegArray.UpperBound(1), 8) = True And xaMultiLegArray(grdMultLegView.Bookmark, 8) = True Then
            sIVISInfo = sDriverDomicile & "," & sIVISInfo & "," & sUserName & "," & Trim$(JobSysNrOid) & "," & Trim$(pInst.PDInfo.SegmentSystemNumberOID)
        'Else
        '    sIVISInfo = sDriverDomicile & "," & sIVISInfo & "," & sUserName & "," & Trim$(JobSysNrOid) & "," & Trim$(pInst.PDInfo.SegmentSystemNumberOID)
        'End If
       
        If g_iDebug = 13 Then
         Call InfoLog("Passed to IVIS - " & sIVISInfo)
        End If
        Call oIVISEditor.startIVISEditor("", "", sIVISInfo)
       
        'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
       
        Set oIVISEditor = Nothing
        Set oWSInfo = Nothing
    End If
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF6AddUpdateLeg_Click() - End")
    End If
Exit Sub
ErrorHandler:
    glErrNum = 400
    gsError = "mnuF6AddUpdateLeg_Click"
    Call oProc.update_error_object(Me, gsError)
    Call oErrorObject.error_routine(oEventlog.FeederShell, _
                                    Err.Number, _
                                    Err.Description & " - Make sure the IVIS client is installed on workstation.", _
                                    oProc, _
                                    ERROR_MSG, _
                                    FEEDER_DISPATCH_DRIVER, , "Make sure the IVIS client is installed on workstation.")
    Set oEventlog = Nothing
End Sub
 
Private Sub mnuF7MultipleLoad_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF7MultipleLoad_Click() - Begin")
    End If
    DoDspMultiLd
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuF7MultipleLoad_Click() - End")
    End If
End Sub
 
Private Sub mnuHelp_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuHelp_Click() - Begin")
    End If
   Form_Help Me
   If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuHelp_Click() - End")
    End If
End Sub
 
Private Sub mnuNPDLdsDisplay_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuNPDLdsDisplay_Click() - Begin")
    End If
    GetNonPreDScreen txtEloc.Text
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuNPDLdsDisplay_Click() - End")
    End If
End Sub
 
Private Sub mnuPDLdsDisplay_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPDLdsDisplay_Click - Begin")
    End If
    'Team Track #229 - if the job was predispatched, send Job Number.
    'Otherwise, send just the driver name
    If g_bDriverPredispatched = False Then
        GetPreDScreen PDDISPLAY_BOTH, txtEloc.Text, txtDriverName.Text, txtJobNr.Text
    Else
        GetPreDScreen CLVT_DRIVER, txtEloc.Text, txtDriverName.Text, txtJobNr.Text
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPDLdsDisplay_Click - End")
    End If
End Sub
 
Private Sub mnuPrint_To_File_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPrint_To_File_Click() - Begin")
    End If
  Print_records ToFile
  If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPrint_To_File_Click() - End")
    End If
End Sub
 
Private Sub mnuPrint_To_Printer_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPrint_To_Printer_Click() - Begin")
    End If
  Print_to_Printer
  If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPrint_To_Printer_Click() - End")
    End If
End Sub
 
Private Sub mnuPrint_To_Spreadsheet_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPrint_To_Spreadsheet_Click() - Begin")
    End If
  Print_records ToExcelSpreadSheet
  If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuPrint_To_Spreadsheet_Click() - End")
    End If
End Sub
 
Private Sub mnuReset_Click()
Dim i As Integer
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuReset_Click() - Begin")
    End If
    If bExitScreen = True Then
      Unload Me  'QS 1/31/05
    Else
      If ctlWarningPopup.WarningVisible Then
         ctlWarningPopup.ClearWarning
      End If
      Call ResetForm
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuReset_Click() - End")
    End If
End Sub
 
Private Sub mnuSchByLds_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuSchByLds_Click() - Begin")
    End If
    DoDspSchedByLoadName
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuSchByLds_Click() - End")
    End If
End Sub
 
Private Sub mnuStandardToolbar_Click()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuStandardToolbar_Click() - Begin")
    End If
   If tbrStd.Visible = True Then
        tbrStd.Visible = False
        mnuStandardToolbar.Checked = False
        repositionFormControls False
    Else
        tbrStd.Visible = True
        mnuStandardToolbar.Checked = True
        repositionFormControls True
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD mnuStandardToolbar_Click() - End")
    End If
End Sub
 
Private Sub oNotify_NotifyWithMultiBaySearchDataF4(uHFCSLoad As GLOBALDEFS.HFCSLOAD, sDestinationHwnd As Long)
  Dim i As Integer
  Dim iBayPosition As Integer
  Dim oColBay As HFCSYardObject.Bay
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD oNotify_NotifyWithMultiBaySearchDataF4 - Begin.")
    End If
  'if 14 screen already closed, then do not responded to F4
  If OGlobalFormNames Is Nothing Then
      If g_iDebug = 13 Then
         Call InfoLog("frmPDD mnuStandardToolbar_Click() - End - 14 screen already closed")
      End If
      Exit Sub
  End If
 
  Me.Enabled = True
 
  bFromF4F5 = True
  bDatafromMatchingScreen(iBayIndex) = False
  bDuplicateLoad = False
 
  'When arrival load is deleted, then delete the dolly which came with the trailer.
     If bFromGTForm(iBayIndex) And Len(txtDolly.Item(iBayIndex).Text) > 0 Then
        txtDolly.Item(iBayIndex).Text = vbNullString
     End If
    
     bDatafromMatchingScreen(iBayIndex) = False
     bDataFromEloc(iBayIndex) = False
     bFromGTForm(iBayIndex) = False
    
     If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
     End If
    
     
  If bBayCompeleted(iBayIndex) Then
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD oNotify_NotifyWithMultiBaySearchDataF4 - End - bay completed begin.")
       End If
       Exit Sub
    End If
     
      If m_sChildName <> OGlobalFormNames.MultiBaySearchF4 Then
          If g_iDebug = 13 Then
              Call InfoLog("frmPDD oNotify_NotifyWithMultiBaySearchDataF4 - End - child name begin.")
          End If
          Exit Sub
      End If
  If Me.hwnd = sDestinationHwnd Then
   
    If ctlBayNr.LBound <= iBayIndex And ctlBayNr.UBound >= iBayIndex Then
        If Len(ctlBayNr.Item(iBayIndex).Text) > 0 Then
            If ctlBayNr.Item(iBayIndex).Text <> uHFCSLoad.BayNumber Then
                ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
                For i = 1 To 3
                    If pInst.raLockedBay(i) = ctlBayNr.Item(iBayIndex).Text Then
                        ' Unlock the Bay
                        pInst.raLockedBay(i) = 0
                       
                        iBayPosition = 1
                        For Each oColBay In g_colBays
                            If ctlBayNr(iBayIndex).Text = oColBay.Number Then
                               g_colBays.Remove iBayPosition
                            End If
                            iBayPosition = iBayPosition + 1
                        Next
                       
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
    'Lock    this bay for predispatch
    ProcessRetData uHFCSLoad
   
    If bShiftLdInfo Then
       For i = 1 To pInst.PDInfo.hPDHll.Count - (iBayIndex + 1)
           Call ShiftLdInfo
       Next i
    End If
    LockBayForPredispatch (iBayIndex)
  End If
 
 ' m_sChildName = vbNullString
 
  ctlBayNr.Item(iBayIndex).SelStart = 0
  ctlBayNr.Item(iBayIndex).SelLength = Len(ctlBayNr.Item(iBayIndex).Text)
 
'  For i = 1 To 200
'   DoEvents - Removed 9-16-14 to resolve flow issues with the screen
'  Next i
   
  If g_iDebug = 13 Then
      Call InfoLog("frmPDD oNotify_NotifyWithMultiBaySearchDataF4 - End.")
  End If
End Sub
 
Private Sub oNotify_NotifyWithPredispatchData(uHFCSLoad As GLOBALDEFS.HFCSLOAD, sDestinationHwnd As Long)
  Dim i As Integer
 
  If g_iDebug = 13 Then
      Call InfoLog("frmPDD oNotify_NotifyWithPredispatchData - Begin.")
  End If
  'if 890 screen already closed, then do not responded to F8
      If OGlobalFormNames Is Nothing Then
          If g_iDebug = 13 Then
              Call InfoLog("frmPDD oNotify_NotifyWithPredispatchData oGlobalFormNames - End")
          End If
          Exit Sub
      End If
 
  Me.Enabled = True
 
  If bBayCompeleted(iBayIndex) Then
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD oNotify_NotifyWithPredispatchData Bay Completed - End")
       End If
       Exit Sub
  End If
  bFromF4F5 = True
  bDatafromMatchingScreen(iBayIndex) = False
  bDuplicateLoad = False
 
  'When arrival load is deleted, then delete the dolly which came with the trailer.
  If bFromGTForm(iBayIndex) And Len(txtDolly.Item(iBayIndex).Text) > 0 Then
     txtDolly.Item(iBayIndex).Text = vbNullString
  End If
   
  bDatafromMatchingScreen(iBayIndex) = False
  bDataFromEloc(iBayIndex) = False
  bFromGTForm(iBayIndex) = False
   
  If ctlWarningPopup.WarningVisible Then
     ctlWarningPopup.ClearWarning
  End If
    
  If bIntxtDriverNa Then
     If g_iDebug = 13 Then
         Call InfoLog("frmPDD oNotify_NotifyWithPredispatchData bIntxtDriverNa - End")
     End If
     Exit Sub
  End If
 
  If m_sChildName <> OGlobalFormNames.PreDispatchedLoads And m_sChildName <> OGlobalFormNames.NonPreDispatchedLoads Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD oNotify_NotifyWithPredispatchData check ChildName - End")
        End If
        Exit Sub
  End If
     
  If Me.hwnd = sDestinationHwnd Then
   
    ProcessRetData uHFCSLoad
   
    'Lock this bay for predispatch
    LockBayForPredispatch (iBayIndex)
End If
 m_sChildName = vbNullString
ctlBayNr.Item(iBayIndex).SelStart = 0
ctlBayNr.Item(iBayIndex).SelLength = Len(ctlBayNr.Item(iBayIndex).Text)
' For i = 1 To 200
'   'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
' Next i
If g_iDebug = 13 Then
     Call InfoLog("frmPDD oNotify_NotifyWithPredispatchData - End")
End If
End Sub
 
Private Sub oNotify_NotifyWithQuickSearchData(uHFCSLoad As GLOBALDEFS.HFCSLOAD, sDestinationHwnd As Long)
  Dim i As Integer
  On Error GoTo Error_Handler
  If g_iDebug = 13 Then
      Call InfoLog("frmPDD oNotify_NotifyWithQuickSearchData - Begin.")
  End If
  'if 14 screen already closed, then do not responded to F4
  If OGlobalFormNames Is Nothing Then
      If g_iDebug = 13 Then
          Call InfoLog("frmPDD oNotify_NotifyWithQuickSearchData oGlobalFormNames - End")
      End If
      Exit Sub
  End If
  Me.Enabled = True
  If bBayCompeleted(iBayIndex) Then
      If g_iDebug = 13 Then
          Call InfoLog("frmPDD oNotify_NotifyWithQuickSearchData Bay Completed - End")
      End If
      Exit Sub
  End If
 
  bFromF4F5 = True
  bDatafromMatchingScreen(iBayIndex) = False
 bDuplicateLoad = False
 
  'When arrival load is deleted, then delete the dolly which came with the trailer.
  If bFromGTForm(iBayIndex) And Len(txtDolly.Item(iBayIndex).Text) > 0 Then
      txtDolly.Item(iBayIndex).Text = vbNullString
  End If
   
  bDatafromMatchingScreen(iBayIndex) = False
  bDataFromEloc(iBayIndex) = False
  bFromGTForm(iBayIndex) = False
   
  If ctlWarningPopup.WarningVisible Then
     ctlWarningPopup.ClearWarning
  End If
    
  If m_sChildName <> OGlobalFormNames.QuickSearchF5 Then
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD oNotify_NotifyWithQuickSearchData QuickSearchF5 - End")
       End If
       Exit Sub
  End If
   
  If Me.hwnd = sDestinationHwnd Then
     ProcessRetData uHFCSLoad
     'Lock this bay for predispatch
     LockBayForPredispatch (iBayIndex)
  End If
   
 ' m_sChildName = vbNullString
  ctlBayNr.Item(iBayIndex).SelStart = 0
  ctlBayNr.Item(iBayIndex).SelLength = Len(ctlBayNr.Item(iBayIndex).Text)
   
'  For i = 1 To 200
'     'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
'  Next i
  If g_iDebug = 13 Then
     Call InfoLog("frmPDD oNotify_NotifyWithQuickSearchData - End")
  End If
  Exit Sub
Error_Handler:
glErrNum = 400
gsError = "oNotify_NotifyWithQuickSearchData"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                           IIf(Err.Number <> 0, Err.Number, glErrNum), _
                           Err.Description & " Module:" & gsError, _
                           oProc, _
                           ERROR_MSG, _
                           FEEDER_DISPATCH_DRIVER) Then
Set oEventlog = Nothing
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'Unload Me
End If
End Sub
 
Private Sub oNotify_NotifyWithTrailerLoadData(uHFCSLoad As GLOBALDEFS.HFCSLOAD)
    If g_iDebug = 13 Then
     Call InfoLog("frmPDD oNotify_NotifyWithTrailerLoadData - Begin")
    End If
    If OGlobalFormNames Is Nothing Then Exit Sub
    ProcessRetData uHFCSLoad
    If g_iDebug = 13 Then
     Call InfoLog("frmPDD oNotify_NotifyWithTrailerLoadData - End")
    End If
End Sub
 
Private Sub oshell_NotifyShellOfApplicationEnd(sScreenName As String)
   Dim rsOnPropInfo As ADODB.Recordset
   Dim tMsg As MSG
 
   Do While PeekMessage(tMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
        'do nothing, just clear out buffer
   Loop
   
   'Send info to ArrivalTractorOnly. PreDipatched True/False and PDP end = True.
   On Error GoTo ErrorHandler
  
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD oshell_NotifyShellOfApplicationEnd - Begin")
'      If Not ctlControls Is Nothing Then
'          Call InfoLog("frmPDD oshell_NotifyShellOfApplicationEnd " & ctlControls.Name & " " & Me.Enabled)
'      End If
   End If
  
   If Not OGlobalFormNames Is Nothing Then
        Select Case sScreenName
            Case OGlobalFormNames.TrailerLoadMessages
                Me.Enabled = True
               
                'prevent from putting focus back on the first bay number everytime
                If ctlControls.Enabled = True Then ctlControls.SetFocus
'                If Len(ctlBayNr.Item(iBayIndex).Text) > 0 Then
'                    ctlBayNr.Item(iBayIndex).SetFocus
'                Else
'                    Exit Sub
'                End If
'
                If Not oClsDb Is Nothing Then
                    'refresh the load and trailer message IDs after leaving the trailer/load message screen.
                    'If a trailer/load message was added, highlight the columns as needed.  This can be done
                    'by getting the information from TFOPTPR.
                    Set rsOnPropInfo = oClsDb.GetOnPropertyLoadInfo(pInst.PDInfo.hPDHll.Item(iBayIndex + 1).lCurFopTlrEntity)
                    If Not rsOnPropInfo Is Nothing Then
                        If Not rsOnPropInfo.EOF Then
                            rsOnPropInfo.MoveFirst
                           
                            If Not IsNull(rsOnPropInfo.Fields("LD_MSG_SEQ_NR").Value) Then
                                pInst.PDInfo.hPDHll.Item(iBayIndex + 1).lLdMsgId = rsOnPropInfo.Fields("LD_MSG_SEQ_NR").Value
                            Else
                                pInst.PDInfo.hPDHll.Item(iBayIndex + 1).lLdMsgId = 0
                            End If
                           
                            If Not IsNull(rsOnPropInfo.Fields("TLR_MSG_SEQ_NR").Value) Then
                                pInst.PDInfo.hPDHll.Item(iBayIndex + 1).lTlrMsgId = rsOnPropInfo.Fields("TLR_MSG_SEQ_NR").Value
                            Else
                                pInst.PDInfo.hPDHll.Item(iBayIndex + 1).lTlrMsgId = 0
                            End If
                            'redisplay the load information, highlighting any necessary columns
                            Call DisplayLoadInfo(pInst.PDInfo.hPDHll(iBayIndex + 1), iBayIndex)
                            sLoadMsg = vbNullString
                            sTrailerMsg = vbNullString
                        End If
                    End If
                End If
            Case OGlobalFormNames.ScheduledInformation, _
                 OGlobalFormNames.MultiBaySearchF4, _
                 OGlobalFormNames.SchedulesByLoadNameOutbound, _
                 OGlobalFormNames.QuickSearchF5, _
                 OGlobalFormNames.NonPreDispatchedLoads, _
                 OGlobalFormNames.EditDollies, _
                 OGlobalFormNames.PreDispatchedLoads, _
                 OGlobalFormNames.YardSearch
                
                 Me.Enabled = True
                
                 'set focus to item. When enabling the screen after it has
                 'been disabled, the focus is either going to the job number
                 'or bay number.  We do not want to change focus when the screen
                 'is enabled
'                 If Len(ctlBayNr.Item(iBayIndex).Text) > 0 Then
'                    ctlBayNr.Item(iBayIndex).SetFocus
'                 Else
'                 If Len(ctlBayNr.Item(iBayIndex).Text) > 0 Then
'                    ctlBayNr.Item(iBayIndex).SetFocus
'                 Else
                  If Not ctlControls Is Nothing Then
                      If ctlControls.Enabled = True Then
                        Me.Tag = ctlControls.Name
                        ctlControls.SetFocus
                      End If
                  End If
'                 End If
        End Select
    End If
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD oshell_NotifyShellOfApplicationEnd - End")
   End If
Exit Sub
ErrorHandler:
  gsError = "oshell_NotifyShellOfApplicationEnd"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
        'Unload Me
    End If
End Sub
 
Private Sub oshell_NotifyShellofApplicationStart(lHwnd As Long, sScreenName As String)
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD oshell_NotifyShellOfApplicationStart - Begin")
   End If
    If Not OGlobalFormNames Is Nothing Then
        Select Case sScreenName
            Case OGlobalFormNames.ScheduledInformation, _
                 OGlobalFormNames.TrailerLoadMessages, _
                 OGlobalFormNames.MultiBaySearchF4, _
                 OGlobalFormNames.SchedulesByLoadNameOutbound, _
                 OGlobalFormNames.QuickSearchF5, _
                 OGlobalFormNames.NonPreDispatchedLoads, _
                 OGlobalFormNames.EditDollies, _
                 OGlobalFormNames.PreDispatchedLoads, _
                 OGlobalFormNames.YardSearch
               
        '        If (m_sChildName = sScreenName) Then
                 Set ctlControls = Me.ActiveControl
                 Me.Enabled = False
        '        End If
       
        End Select
    End If
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD oshell_NotifyShellOfApplicationStart - End")
   End If
End Sub
 
Private Sub SendCancelPredispatch()
    Dim oCancelLoads(2) As HFCSLoadObject.HFCSLOAD
    Dim oCancelLoad As HFCSLoadObject.HFCSLOAD
    Dim iCancelCount As Integer
    Dim bCancelPredispatch As Boolean
    Dim oWsInformation As HFCSWSInformation.clsWSInfo
    Dim oSendPredispatchtoIVIS As PredispatchObject.PreDispatch
    Dim i As Integer
    Dim iLoadCount As Integer
    Dim iLoadsNotCanceled As Integer
    Dim eLoadUpdate As enumLoadUpdate
   
    Set oWsInformation = New HFCSWSInformation.clsWSInfo
   
    On Error GoTo Error_Handler
       
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD SendCancelPredispatch - Begin")
    End If
   
    iLoadCount = 0
    iCancelCount = 0
 
    iLoadCount = pInst.PDInfo.hPDHll.Count
    For i = 1 To iLoadCount
        If Not IsNull(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).AssignedPDOID) Then
            If Trim$(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).AssignedPDOID) = Trim$(pInst.PDInfo.SegmentSystemNumberOID) Then
                Set oCancelLoad = New HFCSLoadObject.HFCSLOAD
                oCancelLoad.SpecifyConnection oClsDb.DBClass
                Call oCancelLoad.GetLoadByEntityKey(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lCurFopTlrEntity, False)
                oCancelLoad.BayNumber = ctlBayNr(i - 1).Text
                oCancelLoad.OutboundFRDTrailerGenNR = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lCurForTlrEntity
                oCancelLoad.OriginCountryCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szOriginCn
                oCancelLoad.OriginName = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szOrigin
                oCancelLoad.OriginCountryCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szOriginCn
                oCancelLoad.DestinationName = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szDestin
                oCancelLoad.DestinationSortTypeCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szDestinSrt
                oCancelLoad.SequenceNumber = Format$(CStr(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).iSequenceNr), "00")
                oCancelLoad.DestinationCountryCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szDestinCn
                If pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).dtLdCrtDt <> 0 Then
                    oCancelLoad.CreateDate = FmtDbDate(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).dtLdCrtDt)
                End If
                oCancelLoad.EqpDirectionTypeCode = "O"
                oCancelLoad.SchTlrEqpGenNr = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lSchTlrEqpGenNr
                'HFCS_INDICATOR_TRUE
                oCancelLoad.PredispatchJob = txtJobNr.Text
                If Len(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szCurPDJobDomCn) > 0 Then
                   oCancelLoad.OutboundActSLOCCountryCD = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szCurPDJobDomCn
                Else
                   oCancelLoad.OutboundActSLOCCountryCD = pInst.PDInfo.szPDJobDomCn
                End If
                If Len(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szCurPDJobDom) > 0 Then
                    oCancelLoad.OutboundActSLOC = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szCurPDJobDom
                Else
                    oCancelLoad.OutboundActSLOC = pInst.PDInfo.szPDJobDomSlic
                End If
                oCancelLoad.PredispatchDate = Date
                oCancelLoad.PredispatchTime = Format(Now, "HH:MM:SS")
                oCancelLoad.OriginSortTypeCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szOriginSrt
               
                'KG - WR00590, if ELOC in structure is not set, use ELOC Country variable
                'from PDInfo
               ' If Len(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szElocCn) > 0 Then
               '     oCancelLoad.PredispatchELOCCnyCd = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szElocCn
               ' Else
                oCancelLoad.PredispatchELOCCnyCd = pInst.PDInfo.szPDElocCn
               ' End If
                '*** end WR00590
               
                oCancelLoad.PredispatchedSegmentSystemOID = pInst.PDInfo.SegmentSystemNumberOID
                oCancelLoad.JobNumberSystemOID = pInst.PDInfo.JobSystemNumberOID
                oCancelLoad.PredispatchELOC = Left$(txtEloc.Text, 5)
                oCancelLoad.LoadMsgSeqNumber = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lLdMsgId
                oCancelLoad.TrailerMsgSeqNumber = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lTlrMsgId
                If (Len(Trim$(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szTlrNr)) > 0) Then
                    oCancelLoad.TrailerNumber = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szTlrNr
                End If
                oCancelLoad.TrailerTypeCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szEqpTyp
                If pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szMultLdIr Then
                   oCancelLoad.HasMultipleLoads = "1"
                Else
                   oCancelLoad.HasMultipleLoads = "0"
                End If
                oCancelLoad.PredispatchDriver = txtDriverName.Text
                oCancelLoad.PredispatchTractor = txtTractorNumber.Text
               
                If pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lCurFopTlrEntity <> 0 Then
                 oCancelLoad.EntityKey = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lCurFopTlrEntity
                End If
               
                oCancelLoad.OutboundVehicleEntityKey = pInst.PDInfo.lPDFopVehEntKey  'pInst.PDInfo.hPDHll.Item(1).lSchdVehEntity
                oCancelLoad.DollyKey = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lCurFopDolEntity
                If Len(pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szLdCd) > 0 Then
                    oCancelLoad.LoadCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szLdCd
                End If
                oCancelLoad.Position = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szTlrPosCd
                oCancelLoad.RouteCode = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).szRtId
                oCancelLoad.DriverNotified = "1" 'pInst.PDInfo.szFdrDvrInfNtfIr
                oCancelLoad.PackageCount = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lPieces
                oCancelLoad.PackagePercentage = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).iPercent
                oCancelLoad.DollyName = txtDolly(0).Text
               
                If pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lSchFdrEntKey <> -1 Then
                    oCancelLoad.PreDispatchScheduledEntityKey = pInst.PDInfo.hPDHll.Item(iLoadsNotCanceled + 1).lSchFdrEntKey
                End If
               
                oCancelLoad.OutboundFeederScheduleWeekEndingDate = ctlWkendingDate.Value
                oCancelLoad.OutboundFeederScheduleDOW = cboDow.ListIndex
                'oCancelLoad.OutboundSchedDate = Format(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy")
               
                oCancelLoad.DepartureDate = Format(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy")
                oCancelLoad.DepartureTime = Format(pInst.PDInfo.FeederScheduleEndTime, "hh:mm")
                If i = iBayIndex + 1 Then
                    pInst.PDInfo.hPDHll.Remove (i - iCancelCount)
                    oCancelLoad.Predispatched = False
                   
                    'mve TT00512
                    If IsBlank(oCancelLoad.TrailerNumber) Then
                       eLoadUpdate = oCancelLoad.CancelPredispatchEmpty(0)
                       bPredispatchEmpty(i - 1) = False
                       If eLoadUpdate = LOAD_UPDATE_SUCCESS Then
                        oCancelLoad.SendCancelPredispatchEmptyAuditTrail
                       End If
                    Else
                       eLoadUpdate = oCancelLoad.Update   ' .CancelPredispatch(0)
                       If eLoadUpdate = LOAD_UPDATE_SUCCESS Then
                        oCancelLoad.SendCancelPredispatchtoAuditTrail
                       End If
                    End If
                    
                    If i = 1 Then
                        bBayCompeleted(0) = bBayCompeleted(1)
                        bBayCompeleted(1) = bBayCompeleted(2)
                        bBayCompeleted(2) = False
                        bctlBayValidateComplete(0) = bctlBayValidateComplete(1)
                        bctlBayValidateComplete(1) = bctlBayValidateComplete(2)
                        bctlBayValidateComplete(2) = False
                    ElseIf i = 2 Then
                        bBayCompeleted(1) = bBayCompeleted(2)
                        bBayCompeleted(2) = False
                        bctlBayValidateComplete(1) = bctlBayValidateComplete(2)
                        bctlBayValidateComplete(2) = False
                    Else
                        bBayCompeleted(2) = False
                        bctlBayValidateComplete(2) = False
                    End If
                    iCancelCount = iCancelCount + 1
                Else
                    Set oCancelLoads(iLoadsNotCanceled) = oCancelLoad
                    iLoadsNotCanceled = iLoadsNotCanceled + 1
                End If
       
                Set oCancelLoad = Nothing
            End If
        End If
    Next i
    MsgBox "Cancel predispatch successful"
    If pInst.PDInfo.hPDHll.Count = 0 Then
        Set oCancelLoad = New HFCSLoadObject.HFCSLOAD
        oCancelLoad.SpecifyConnection oClsDb.DBClass
        oCancelLoad.PredispatchedSegmentSystemOID = pInst.PDInfo.SegmentSystemNumberOID
        oCancelLoad.JobNumberSystemOID = pInst.PDInfo.JobSystemNumberOID
        oCancelLoad.OutboundFeederScheduleWeekEndingDate = ctlWkendingDate.Value
        oCancelLoad.OutboundFeederScheduleDOW = cboDow.ListIndex
        oCancelLoad.OutboundActSLOCCountryCD = pInst.PDInfo.szPDJobDomCn
        oCancelLoad.OutboundActSLOC = pInst.PDInfo.szPDJobDomSlic
 
        oCancelLoad.DepartureDate = Format(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy")
        oCancelLoad.DepartureTime = Format(pInst.PDInfo.FeederScheduleEndTime, "hh:mm")
        oCancelLoad.PredispatchELOCCnyCd = pInst.PDInfo.szPDElocCn
        oCancelLoad.PredispatchDriver = txtDriverName.Text
        oCancelLoad.PredispatchTractor = txtTractorNumber.Text
        Set oCancelLoads(iLoadsNotCanceled) = oCancelLoad
        Set oCancelLoad = Nothing
    End If
   
    DisplayLdData pInst.PDInfo.hPDHll
    If pInst.PDInfo.hPDHll.Count = 0 Then
       ctlBayNr(0).Visible = True
       ctlBayNr(0).Enabled = True
       ctlBayNr(0).SetFocus
       ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
       mnuF11CancelPredispatch.Enabled = False
    ElseIf ictlIndex > pInst.PDInfo.hPDHll.Count Then
        If bBayCompeleted(2) = True Then
            ictlIndex = 2
        ElseIf bBayCompeleted(1) = True Then
            ictlIndex = 1
        Else
            ictlIndex = 0
        End If
        ctlBayNr(ictlIndex).Visible = True
        ctlBayNr(ictlIndex).Enabled = True
        ctlBayNr(ictlIndex).SetFocus
       
    Else
        ctlBayNr(ictlIndex).Visible = True
        ctlBayNr(ictlIndex).Enabled = True
        ctlBayNr(ictlIndex).SetFocus
    End If
 
   
    Set oSendPredispatchtoIVIS = New PredispatchObject.PreDispatch
   
    bCancelPredispatch = oSendPredispatchtoIVIS.SendNewPredispatch(oCancelLoads(0), oCancelLoads(1), oCancelLoads(2), Now, ctlWkendingDate.Value, txtJobNr.Text, txtDriverName.Text, CStr(cboDow.ListIndex), txtTractorNumber.Text, txtEloc.Text, sLocalCn, Format$(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy"), Format$(pInst.PDInfo.FeederScheduleEndTime, "hh:mm"), oWsInformation.UserName, oWsInformation.ComputerName)
 
'   oSendPredispatchtoIVIS.Dispose
    Set oSendPredispatchtoIVIS = Nothing
    Set oWsInformation = Nothing
    Set oCancelLoad = Nothing
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD SendCancelPredispatch - End")
    End If
    Exit Sub
Error_Handler:
  gsError = "grdBayDetails_KeyDown"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
End If
End Sub
 
Private Sub txtDolly_GotFocus(Index As Integer)
    On Error GoTo Error_Handler
 
     If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtDolly_GotFocus - Begin - Index: " & Index)
     End If
    
     If (Index <> ictlIndex) Then
         ictlIndex = Index
         iBayIndex = Index
     End If
    
     If bBayCompeleted(Index) = False Then
        If bSetFocusonBay = True Then
            bSetFocusonBay = False
            If bSaveData = False Then
                grdBayDetails(Index).Split = 0
                grdBayDetails(Index).Col = 0
                If ctlBayNr(Index).Enabled = False Then ctlBayNr(Index).Enabled = True
                If ctlBayNr(Index).Visible = False Then ctlBayNr(Index).Visible = True
                ctlBayNr(Index).SetFocus
'                txtDolly(Index).Enabled = False
            End If
          '  If ctlBayNr(Index).Enabled = False Then ctlBayNr(Index).Enabled = True
          '  If ctlBayNr(Index).Visible = False Then ctlBayNr(Index).Visible = True
       '     txtDolly(Index).Enabled = False
         '   ctlBayNr(Index).SetFocus
          '  DoEvents
        ElseIf bGridGotFocus = True And ctlWarningPopup.Visible = True Then
            bGridGotFocus = False
            If grdBayDetails(Index).Enabled = False Then grdBayDetails(Index).Enabled = True
            grdBayDetails(Index).SetFocus
        End If
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD txtDolly_GotFocus - End - If bBayCompeleted(Index) = False - Index: " & Index)
        End If
        Exit Sub
    End If
 
bLoadRoutingError = False
    txtDolly(Index).BackColor = vbGreen
'<WR01024><ADID:-dmx7plm> - Start
    If bClickedF4 = True Then mCtlbayfirst = False
    If bClickedF5 = True Then mCtlbayfirst = False
    If Not IsUserShiftTabbing Then
        If Index >= iBayIndex Then
            iBayIndex = Index
            ictlIndex = Index
        End If
    End If
'<WR01024><ADID:-dmx7plm> - End
    If bPredispatchEmpty(Index) = True Then SetGridNormal (Index)
    If Not mCtlbayfirst Then
       If Len(ctlBayNr(Index).Text) = 0 Then
           SetAvailMenuAndTool mat_dolly, False, False
       Else
           SetAvailMenuAndTool mat_dolly, True, False
       End If
    End If
 
    If pInst.PDInfo.hPDHll.Count > Index Then
        If Not IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) And Not IsNull(pInst.PDInfo.SegmentSystemNumberOID) Then
            If Trim$(pInst.PDInfo.SegmentSystemNumberOID) = Trim$(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) Then
                ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, True
                mnuF11CancelPredispatch.Enabled = True
            Else
               ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
                mnuF11CancelPredispatch.Enabled = False
            End If
        Else
            ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
            mnuF11CancelPredispatch.Enabled = False
        End If
    Else
        ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
        mnuF11CancelPredispatch.Enabled = False
    End If
 
    mCtlbayfirst = False
    iDollyIndex = Index
    chkDriverNotified.Enabled = True
    lblDeleteRecord.Visible = False
 
    bExitScreen = False
 
'    If g_iDebug = 13 Then
'      Call InfoLog("frmPDD txtDolly_GotFocus end")
'    End If
 
'        With ctlToolbarManager
'            .buttonEnabled BT_F9_DOLLIES_ON_PROP, mnuDolliesDisplay.Enabled
'            .buttonEnabled BT_F10_ACCEPT, mnuF10Accept.Enabled
'        End With
'        mnuDolliesDisplay.Enabled = True
'        mnuF10Accept.Enabled = True
    If g_iDebug = 13 Then
            Call InfoLog("frmPDD txtDolly_GotFocus - End - Index: " & Index)
    End If
    Exit Sub
Error_Handler:
  gsError = "txtDolly_GotFocus"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Sub
 
Private Sub txtDolly_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
   
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDolly_KeyDown - Begin - Index: " & Index)
    End If
    If gbLoadPreDispatched = True Then
        KeyCode = 0
        If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtDolly_KeyDown LoadPredispatched End - Index: " & Index)
        End If
        Exit Sub
    End If
    If Len(ctlBayNr.Item(Index).Text) = 0 And Not bPredispatchEmpty(Index) Then
       KeyCode = 0
       If g_iDebug = 13 Then
         Call InfoLog("frmPDD txtDolly_KeyDown no bay number for index " & Index & " - End")
       End If
       Exit Sub
    End If
    bDollyValidated(Index) = False
    If KeyCode = vbKeyTab Then
       mbtab = True
    End If
    If KeyCode = vbKeyF10 And Shift = 0 And mnuF10Accept.Enabled Then
        DoSaveData
    End If
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDolly_KeyDown - End - Index: " & Index)
    End If
    Exit Sub
ErrorHandler:
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
      '  Unload Me
    End If
End Sub
 
Private Sub txtDolly_LostFocus(Index As Integer)
   On Error GoTo ErrorHandler
  
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDolly_LostFocus - Begin - Index: " & Index)
   End If
   If bBayCompeleted(Index) = False Then
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtDolly_LostFocus If bBayCompeleted(Index) = False - End - Index: " & Index)
      End If
      Exit Sub
   End If
   txtDolly(Index).BackColor = vbWhite
   If IsUserShiftTabbing Then
      bSaveData = False
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtDolly_LostFocus UserShiftTabbing - End - Index: " & Index)
      End If
      Exit Sub
   End If
   
   If bSaveData And GetKeyState(vbKeyTab) < 0 And mnuF10Accept.Enabled = True Then
      DoSaveData
   End If
   If g_iDebug = 13 Then
     Call InfoLog("frmPDD txtDolly_LostFocus - End - Index: " & Index)
   End If
Exit Sub
 
ErrorHandler:
    glErrNum = 400
    gsError = "txtDolly_LostFocus"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
    End If
End Sub
 
Private Sub txtDolly_Validate(Index As Integer, Cancel As Boolean)
Dim i As Integer
Dim k As Integer
Dim bFnd As Boolean
Dim lEntKey As Long
Dim sPreDollyNr As String
 
    On Error GoTo ErrorExit
 
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDolly_Validate - Begin - Index: " & Index)
    End If
   
    bDollyValidated(Index) = False
   
    If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
    End If
   
    If pInst.PDInfo.hPDHll.Count - 1 < Index Then
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtDolly_Validate count < index - End - Index: " & Index)
       End If
       Exit Sub
    End If
 
      'Check duplicate dolly
    If pInst.PDInfo.hPDHll.Count = 2 Then
       If txtDolly(0).Text = txtDolly(1).Text And Len(txtDolly(0).Text) > 0 And Len(txtDolly(1).Text) > 0 Then
          ctlWarningPopup.ShowWarning txtDolly(1).hwnd, LoadResString(sDuplicateDolly), 2000
          Cancel = True
          txtDolly(1).SelStart = 0
          txtDolly(1).SelLength = Len(txtDolly(1).Text)
          If g_iDebug = 13 Then
              Call InfoLog("frmPDD txtDolly_Validate Dolly0 = Dolly1 - count 2 - End - Index: " & Index)
          End If
          Exit Sub
       End If
    ElseIf pInst.PDInfo.hPDHll.Count = 3 Then
       If txtDolly(0).Text = txtDolly(2).Text And Len(txtDolly(0).Text) > 0 And Len(txtDolly(2).Text) > 0 Then
          ctlWarningPopup.ShowWarning txtDolly(2).hwnd, LoadResString(sDuplicateDolly), 2000
          txtDolly(2).SelLength = Len(txtDolly(2).Text)
          Cancel = True
          If g_iDebug = 13 Then
              Call InfoLog("frmPDD txtDolly_Validate Dolly0 = Dolly2 - count 3 - End - Index: " & Index)
          End If
          Exit Sub
       End If
      If txtDolly(0).Text = txtDolly(1).Text And Len(txtDolly(0).Text) > 0 And Len(txtDolly(1).Text) > 0 Then
          ctlWarningPopup.ShowWarning txtDolly(2).hwnd, LoadResString(sDuplicateDolly), 2000
          txtDolly(1).SelLength = Len(txtDolly(1).Text)
          Cancel = True
          If g_iDebug = 13 Then
              Call InfoLog("frmPDD txtDolly_Validate Dolly0 = Dolly1 - count 3 - End - Index: " & Index)
          End If
          Exit Sub
       End If
       If txtDolly(1).Text = txtDolly(2).Text And Len(txtDolly(1).Text) > 0 And Len(txtDolly(2).Text) > 0 Then
          ctlWarningPopup.ShowWarning txtDolly(2).hwnd, LoadResString(sDuplicateDolly), 2000
          txtDolly(2).SelLength = Len(txtDolly(2).Text)
          Cancel = True
          If g_iDebug = 13 Then
              Call InfoLog("frmPDD txtDolly_Validate Dolly1 = Dolly2 - count 3 - End - Index: " & Index)
          End If
          Exit Sub
       End If
    End If
   
    If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
    End If
    lblDeleteRecord.Visible = False
    ' Determine which BAY NUMBER field
    ' focus is on.
    If Not IsBlank(txtDolly(Index).Text) Then
       If pInst.bDolDispatched(Index + 1) And _
            pInst.PDInfo.hPDHll(Index + 1).szCurPDDollyNr = txtDolly(Index).Text Then
           ' Already validated and no change, no need to re-validate
           If g_iDebug = 13 Then
              Call InfoLog("frmPDD txtDolly_Validate bDolDispatched - End - Index: " & Index)
           End If
           Exit Sub
       ElseIf bFromGTForm(Index) And Len(txtDolly(Index).Text) > 0 Then
           For k = 1 To 3  'no validate for the arrival dolly
              If txtDolly(Index).Text = raArrivalData(Index + 1).sActDolNr Then
                 bDollyValidated(Index) = True
                 If g_iDebug = 13 Then
                    Call InfoLog("frmPDD txtDolly_Validate bFromGtFrom - End - Index: " & Index)
                 End If
                 Exit Sub
              End If
           Next k
       End If
 
 
        If pInst.ValidateDolly(txtDolly(Index).Text, lEntKey) Then
            bFnd = True
            ' If valid then check to see if it is already been used
            For i = 1 To pInst.PDInfo.hPDHll.Count
                    pInst.PDInfo.hPDHll(Index + 1).lCurFopDolEntity = lEntKey
            Next i
 
            If Not bFnd Then
                ' Increment current Data line
                If pInst.icdl > MAX_DATA_LINES Then
                    ' Post an ACCEPT message to ourself
                    ' We are on the LAST field of the screen
                    AcceptPreDispatch
                    ExitDlg
                End If
                pInst.PDInfo.hPDHll(Index + 1).lCurFopDolEntity = lEntKey
                pInst.PDInfo.hPDHll(Index + 1).szCurPDDollyNr = txtDolly(Index).Text
                pInst.bDolDispatched(Index + 1) = True
            End If
        Else
            Cancel = True
        End If
    End If
 
    If Cancel Then
        txtDolly(Index).SelStart = 0
        txtDolly(Index).SelLength = Len(txtDolly(Index).Text)
    Else
        bDollyValidated(Index) = True
        If Index < 2 And Len(Trim$(txtDolly(Index).Text)) > 0 Then
           chkExtraDolly(Index).Value = 0
           chkExtraDolly(Index).Enabled = False
           chkExtraDolly(Index).TabStop = False
           ctlBayNr(Index + 1).Visible = True
           ctlBayNr(Index + 1).Enabled = True
           ctlBayNr(Index + 1).TabStop = True
           ctlBayNr(Index + 1).SetFocus
        ElseIf Index = 2 Then 'And Len(Trim$(txtDolly(Index).Text)) > 0 Then
           bSaveData = True
        End If
    End If
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDolly_Validate - End - Index: " & Index)
    End If
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "txtDolly_Validate"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                    IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                    Err.Description & " Module:" & gsError, _
                                    oProc, _
                                    ERROR_MSG, _
                                    FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
End Sub
 
Private Sub txtDriverName_GotFocus()
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_GotFocus - Begin")
    End If
   txtDriverName.BackColor = vbGreen
   grdMultLegView.Styles(5).BackColor = vbYellow
   bExitScreen = False
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_GotFocus - End")
   End If
End Sub
 
Private Sub txtDriverName_KeyPress(KeyAscii As Integer)
  If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_KeyPress - Begin")
  End If
  If (KeyAscii = oKeyDefs.Enter_Key) Then
    EnterAsTab KeyAscii
    KeyAscii = KeyToUpperCase(KeyAscii)
   Exit Sub
  End If
  KeyAscii = KeyToUpperCase(KeyAscii)
  If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_KeyPress - End")
  End If
End Sub
 
Private Sub txtDriverName_LostFocus()
   On Error GoTo ErrorHandler
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_LostFocus - Begin")
   End If
   If bOpenPreScreen Then
    'Team Track #177 - if the job was predispatched, send both ELOC and Driver Name.
    'Otherwise, send just the driver name
     If g_bDriverPredispatched = False Then
         bOpenPreScreen = GetPreDScreen(PDDISPLAY_BOTH, txtEloc.Text, txtDriverName.Text, txtJobNr.Text)
     Else
         bOpenPreScreen = GetPreDScreen(CLVT_DRIVER, txtEloc.Text, txtDriverName.Text, txtJobNr.Text)
     End If
   End If
   bOpenPreScreen = False
  ' TT000394
'   txtTractorNumber.Enabled = True
   grdMultLegView.Enabled = True
   If Me.ActiveControl.Name = "grdMultLegView" And grdMultLegView.Visible Then
        If grdMultLegView.Enabled = False Then grdMultLegView.Enabled = True
        'have to check again, when called from the arrivals screen, and the
        'matching loads form comes up, grdMultLegView.Enabled is false, even though
        'it is set to true above
        If grdMultLegView.Enabled = True Then grdMultLegView.SetFocus
   End If
    '   txtEloc.Enabled = True
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_LostFocus end")
   End If
   txtDriverName.BackColor = vbWhite
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_LostFocus - End")
   End If
   Exit Sub
ErrorHandler:
    gsError = "txtDriverName_LostFocus"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
End Sub
 
Private Sub txtDriverName_Validate(Cancel As Boolean)
    Dim bResult As Boolean
 
    On Error GoTo ErrorExit
   
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD txtDriverName_Validate - Begin")
    End If
 
    If Not txtDriverName.Enabled Then
       txtJobNr.Enabled = True
       txtJobNr.SetFocus
       If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtDriverName_Validate not enabled - End")
       End If
       Exit Sub
    End If
    If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
    End If
 
    'Only validate once
    If txtLastDriverName = txtDriverName.Text And Not IsBlank(txtLastDriverName) Then
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtDriverName_Validate lastdrivername = drivername - End")
        End If
        If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtDriverName_Validate not enabled - End - If txtLastDriverName = txtDriverName.Text And...")
        End If
        Exit Sub
    End If
 
      pInst.PDInfo.szPDDvrNa = Trim$(txtDriverName.Text)
    If IsBlank(pInst.PDInfo.szPDDvrNa) Then
        ctlWarningPopup.ShowWarning txtDriverName.hwnd, LoadResString(sDriverNamecannotbeblank), 2250
        If txtDriverName.Enabled = False Then txtDriverName.Enabled = True
        txtDriverName.SetFocus
        Cancel = True
    Else
        If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtDriverName_Validate Blank Else Clause")
        End If
        ' Search for Predispatch Data On Job Number
'        If Not gbArrJobFound Then   'Per Mike A. for arrivals no need to check the job number.QS 2/23/05
        'TSAT #4662 - moved if statement
        'TSAT #4662 - still want to check for job, but don't stay on driver name if not found.
'        bResult = pInst.SearchPDJob()
'        If Not gbArrJobFound Then
'            Cancel = Not bResult
'        End If
    End If
 
    If Cancel And Not pInst.bArriveFlag Then
        If txtDriverName.Enabled = False Then txtDriverName.Enabled = True
        txtDriverName.SetFocus
        txtDriverName.SelStart = 0
        txtDriverName.SelLength = Len(txtDriverName.Text)
        txtLastDriverName = ""
    Else
        txtLastDriverName = txtDriverName.Text
    End If
    If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtDriverName_Validate not enabled - End")
    End If
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "txtDriverName_Validate"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
End Sub
 
Private Sub txtEloc_Change()
Dim lStart As Long
Dim sClosest As String
Static lLock As Long
Static sLastText As String
  If g_iDebug = 13 Then
     Call InfoLog("frmPDD txtEloc_Change - Begin")
  End If
  ' Lock the text box, and find the closest match
  lLock = 1
  sClosest = oClsDb.ClosestSLIC(txtEloc.Text, sLocalCn)
 
  If Len(sClosest) > 0 Then
    lStart = txtEloc.SelStart
    txtEloc.Text = sClosest
    txtEloc.SelStart = lStart
    txtEloc.SelLength = Len(txtEloc.Text) - lStart
  End If
  sLastText = Left$(txtEloc.Text, txtEloc.SelStart)
  lLock = 0
  If g_iDebug = 13 Then
    Call InfoLog("frmPDD txtEloc_Change - End")
  End If
End Sub
 
Private Sub txtEloc_GotFocus()
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtEloc_GotFocus - Begin")
    End If
    chkDriverNotified.Enabled = False
    bExitScreen = False
    TextEnd.Enabled = True
    TextEnd.TabStop = True
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtEloc_GotFocus - End")
    End If
End Sub
 
Private Sub txtEloc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vCtrlDown, vShiftDown As Variant
 
  If g_iDebug = 13 Then
     Call InfoLog("frmPDD txtEloc_KeyDown - Begin")
  End If
 
Select Case KeyCode
  Case vbKeyBack          'User back tabbed off of the SLOC field
    If txtEloc.SelStart <= MAX_SLIC_LENGTH Then
      ' If SLIC entered and backspace hit, delete everything
      txtEloc.Text = vbNullString
    End If
 
  End Select
  If g_iDebug = 13 Then
     Call InfoLog("frmPDD txtEloc_KeyDown - End")
  End If
 
End Sub
 
Private Sub txtEloc_KeyPress(KeyAscii As Integer)
If g_iDebug = 13 Then
     Call InfoLog("frmPDD txtEloc_KeyPress - Begin")
End If
If (KeyAscii = oKeyDefs.Enter_Key) Then
        EnterAsTab KeyAscii
        KeyAscii = KeyToUpperCase(KeyAscii)
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD txtEloc_KeyPress - End - If (KeyAscii = oKeyDefs.Enter_Key)")
        End If
        Exit Sub
    End If
    KeyAscii = KeyToUpperCase(KeyAscii)
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD txtEloc_KeyPress - End")
    End If
End Sub
 
Private Sub txtEloc_LostFocus()
Dim i As Integer
 
On Error GoTo ErrorHandler
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtEloc_LostFocus - Begin")
End If
 
'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
' TT000394 -
If Len(txtTractorNumber.Text) = 0 Then
    txtTractorNumber.Text = "  "
End If
If Len(txtTractorNumber.Text) = 0 And Len(txtJobNr.Text) > 0 And Len(txtDriverName.Text) > 0 Then
'If Len(txtJobNr.Text) > 0 And Len(txtDriverName.Text) > 0 Then
 
   ' TT000394 - HFCS not populating job on 13 screen for the job schedule inserted in IVIS.
'   txtTractorNumber.Enabled = True
'   txtTractorNumber.SetFocus
   If g_iDebug = 13 Then
       Call InfoLog("frmPDD txtEloc_LostFocus enable tractor number - End")
   End If
   Exit Sub
End If
 
'TT#3330 - app2dmw - 3-25-2013 - start
'If Not IsBlank(txtJobNr.Text) And Me.ActiveControl Is TextEnd Then
If Not IsBlank(txtJobNr.Text) And Me.ActiveControl Is TextEnd And (jobLegsUsedUp = False) Then
 
    If (frmBayData.Enabled = False) Then
        frmBayData.Enabled = True
    End If
'TT#3330 - app2dmw - 3-25-2013 - end
   
   'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
    'Only enable the first bay, not all of them.
    ctlBayNr.Item(0).Enabled = True
    ctlBayNr.Item(0).TabStop = True
    ' TT000394
    ctlBayNr.Item(0).Visible = True
    ctlBayNr.Item(0).SetFocus
   
    
    
   If bMatchingForm Then
      If frmMatchingLds.grdLoads.Enabled = False Then frmMatchingLds.grdLoads.Enabled = True
      frmMatchingLds.grdLoads.SetFocus
      If g_iDebug = 13 Then
         Call InfoLog("frmPDD txtEloc_LostFocus grid set focus - End")
      End If
      Exit Sub
   End If
  
'   For i = 0 To 200
'      'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
'   Next i
  
   cboDow.Enabled = False
   ctlWkendingDate.Enabled = False
   chkDriverNotified.Enabled = False
 
   txtJobNr.Enabled = False
   txtDriverName.Enabled = False
   txtTractorNumber.Enabled = False
   txtEloc.Enabled = False
   grdMultLegView.Enabled = False
  
   If iMyRc = 21 Then
      ctlBayNr.Item(0).Visible = True
      ctlBayNr.Item(0).Enabled = True
      ctlBayNr.Item(0).SetFocus
   ElseIf iMyRc = MBID_CANCEL Then
       txtJobNr.Enabled = True
       txtJobNr.SetFocus
   End If
 
ElseIf Not Not Not IsBlank(txtJobNr.Text) And Not (Me.ActiveControl Is TextEnd) And Not pInst.bArriveFlag And Not bMatchingForm Then
  ' txtTractorNumber.Enabled = True
  ' txtTractorNumber.SetFocus
ElseIf iMyRc = MBID_CANCEL Then
   txtJobNr.Enabled = True
   txtJobNr.SetFocus
End If
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtEloc_LostFocus - End")
End If
    Exit Sub
ErrorHandler:
    glErrNum = 400
    gsError = "txtEloc_LostFocus"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
    End If
End Sub
 
Private Sub txtEloc_Validate(Cancel As Boolean)
Dim sSab As String
Dim sCny As String
Dim lResult As Long
Dim iRc As Integer
 
    On Error GoTo ErrorExit
  
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD txtEloc_Validate - Begin")
    End If
   
    If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
    End If
   
    If Not txtEloc.Enabled Then
       If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtEloc_Validate not txtEloc enabled - End")
       End If
       Exit Sub
    End If
   
    If IsUserShiftTabbing And Not pInst.bArriveFlag And Not IsBlank(txtEloc.Text) Then
       If txtTractorNumber.Enabled = False Then txtTractorNumber.Enabled = True
       txtTractorNumber.SetFocus
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtEloc_Validate IsUserShiftTabbing - End")
       End If
       Exit Sub
    End If
   
 
    'The Eloc has been validated  QS 7/27/05
    If txtLastEloc = txtEloc.Text And Not IsBlank(txtLastEloc) And IsUserShiftTabbing Then
        'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtEloc_Validate lastEloc = txtEloc - End")
       End If
       Exit Sub
    End If
   
    ' Gather data from ELOC Field and then Validate
    SplitSabCny txtEloc.Text, sSab, sCny, pInst.PDInfo.szCurrentCny
    pInst.PDInfo.szPDEloc = sSab
    pInst.PDInfo.szPDElocCn = sCny
    txtEloc.Text = MergeSabCny(sSab, sCny, pInst.PDInfo.szCurrentCny)
 
 
    ' Copy ELOC and CN into DB INPUT structure
    If IsBlank(pInst.PDInfo.szPDEloc) Then
        txtEloc.Text = ""
        Cancel = True
        ctlWarningPopup.ShowWarning txtEloc.hwnd, LoadResString(sEloccannotbeblank), 2000
    ' Only Validate ELOC and CN if it had been
    ' changed from what came in from the Schedules
    ElseIf pInst.PDInfo.szSchdEloc <> pInst.PDInfo.szPDEloc Then
        If Not F9802VerifySLC_CN(pInst.PDInfo.szPDEloc, pInst.PDInfo.szPDElocCn) And Not _
           (Trim$(txtEloc.Text) = "DROP" Or Trim$(txtEloc.Text) = "0000" Or Trim$(txtEloc.Text) = "OWZ" _
           Or Trim$(txtEloc.Text) = "TRM.") Then
            ' INVALID SlicCn.  Clear Fields
            txtEloc.Text = ""
            Cancel = True
            ctlWarningPopup.ShowWarning txtEloc.hwnd, LoadResString(sElocisnotvalid), 2000
        Else
            If pInst.bJobNumberFound Then
               ' Search for a scheduled movement
                ' that matches the newly entered ELOC
                lResult = pInst.SchdInfoForEloc()
                If lResult = PD_NODATA Then
                    ' If none can be found, formulate message
                    iRc = pdMsgBox(LoadResString(sNoLoadsforJob) & Space(1) & pInst.PDInfo.szPDJobNr & vbCr & _
                                  LoadResString(sInFeederScheduleforELOC) & pInst.PDInfo.szPDEloc & "/" & pInst.PDInfo.szPDElocCn & vbCr & _
                                  LoadResString(sContinue), _
                                  vbExclamation + vbYesNo, _
                                  LoadResString(sNoLoadsForJobWARNING))
                    If iRc <> vbYes Then
                        Cancel = True
                    Else
                        ' user wants to continue
                        ' clear data retrieved for old eloc
                        Do While pInst.PDInfo.hPDHll.Count > 0
                            pInst.PDInfo.hPDHll.Remove 1
                        Loop
                        mbElocValid = True
                    End If
                ElseIf lResult = PD_SUCCESS Then
                    If pInst.bJobNumberFound Then
                        mbElocValid = pInst.RetrieveSchdMovement(pInst.PDInfo.szPDEloc, pInst.PDInfo.szPDElocCn)
                        If mbElocValid = False Then
                            If pInst.PDInfo.hPDHll.Count > 0 Then
                                mbElocValid = True
                                Call DisplayLdData(pInst.PDInfo.hPDHll)
                            End If
                        End If
                   End If
                ElseIf lResult = PD_FAILURE Then
                    SendKeys ALT_F4_KEY
                    Cancel = True
                End If
            Else ' bjobnumber not found
                ' If none can be found, formulate message
                iRc = pdMsgBox(LoadResString(sNoLoadsforJob) & Space(1) & pInst.PDInfo.szPDJobNr & vbCr & _
                              LoadResString(sInFeederScheduleforELOC) & pInst.PDInfo.szPDEloc & "/" & pInst.PDInfo.szPDElocCn & vbCr & _
                              LoadResString(sContinue), _
                              vbExclamation + vbYesNo, _
                              LoadResString(sNoLoadsForJobWARNING))
                If iRc <> vbYes Then
                    Cancel = True
                Else
                    mbElocValid = True
                End If
            End If
        End If
    Else
        ' Time to Verify and Display Schedule Movement
        ' Validation of all user enterable fields is complete.
        If pInst.bJobNumberFound Then
            mbElocValid = pInst.RetrieveSchdMovement(pInst.PDInfo.szPDEloc, pInst.PDInfo.szPDElocCn)
            If mbElocValid = False Then
                If pInst.PDInfo.hPDHll.Count > 0 Then
                    mbElocValid = True
                    Call DisplayLdData(pInst.PDInfo.hPDHll)
                End If
            End If
        Else
            mbElocValid = True
        End If
    End If
 
    If Not mbElocValid Then
       Cancel = True
    End If
 
 
    If Cancel And Not pInst.bArriveFlag Then
        pInst.PDInfo.szPDEloc = ""
        pInst.PDInfo.szPDElocCn = "" '
       
        If iMyRc = UMBID_DELETE Then
           ctlBayNr.Item(0).Visible = True
           ctlBayNr.Item(0).Enabled = True
           ctlBayNr.Item(0).SetFocus
        End If
 
        txtEloc.SelStart = 0
        txtEloc.SelLength = Len(txtEloc.Text)
        txtLastEloc = ""
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtEloc_Validate bArriveFlag - End")
        End If
        Exit Sub
   
    ElseIf Cancel And pInst.bArriveFlag Then
        txtJobNr.Enabled = False
        cboDow.Enabled = False
        ctlWkendingDate.Enabled = False
        txtLastEloc = txtEloc.Text
    ElseIf Not mbElocValid Then
        txtLastEloc = ""
        txtEloc.SetFocus
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtEloc_Validate mbElocValid - End")
        End If
        Exit Sub
 
    Else
        chkDriverNotified.Enabled = True
        txtLastEloc = txtEloc.Text
    End If
 
    If iMyRc = UMBID_DELETE Or iMyRc = DB_EOF Then   'Added for the condition of removing the load.
       SetAvailMenuAndTool MAT_BASE_DATA0, False, False
    Else  'for no load found  'QS 1/2/05
       If pInst.PDInfo.hPDHll.Count > 0 Then   '5/24
 
            If pInst.PDInfo.hPDHll.Item(1).szMultLdIr = IR_TRUE And (pInst.PDInfo.hPDHll.Item(1).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(1).lLdMsgId > 0) Then
               SetAvailMenuAndTool MAT_BASE_DATA0, True, True
            ElseIf pInst.PDInfo.hPDHll.Item(1).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(1).lLdMsgId > 0 Then
               SetAvailMenuAndTool MAT_BASE_DATA0, True, False
            ElseIf pInst.PDInfo.hPDHll.Item(1).szMultLdIr = IR_TRUE Then
               SetAvailMenuAndTool MAT_BASE_DATA0, False, True
            Else
               SetAvailMenuAndTool MAT_BASE_DATA0, False, False
            End If
        Else
           SetAvailMenuAndTool MAT_BASE_DATA0, False, False
        End If
       'SetAvailMenuAndTool MAT_BASE_DATA0, False, False
 
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD txtEloc_Validate - End")
    End If
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "txtEloc_Validate"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
    End If
End Sub
 
Private Sub txtJobNr_GotFocus()
    On Error GoTo ErrorHandler
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD txtJobNr_GotFocus - Begin")
    End If
    txtJobNr.BackColor = vbGreen
    TextEnd.TabStop = False
    TextEnd.Enabled = False
    If ctlBayNr.LBound < 1 Then
        ctlBayNr.Item(0).TabStop = False
    End If
    ctlBayNr.Item(1).TabStop = False
    ctlBayNr.Item(2).TabStop = False
    txtDolly.Item(0).Enabled = False
    txtDolly.Item(0).BackColor = vbWhite
    txtDolly.Item(1).Enabled = False
    txtDolly.Item(1).BackColor = vbWhite
    txtDolly.Item(2).Enabled = False
    txtDolly.Item(2).BackColor = vbWhite
   
    gbBayReached = False
    SetAvailMenuAndTool MAT_BASE_DATA0, False, False
 
    If Not IsBlank(txtJobNr.Text) Then
       cboDow.Enabled = False
       ctlWkendingDate.Enabled = False   'QS 6/24
       chkDriverNotified.Enabled = False
    End If
 
    If iMyRc = MBID_CANCEL Then
       ctlBayNr(0).Enabled = True
       ctlBayNr.Item(0).Enabled = True
       ClearData = False
    End If
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD txtJobNr_GotFocus - End")
    End If
   
    Exit Sub
ErrorHandler:
    glErrNum = 400
    gsError = "txtJobNr_GotFocus"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
End Sub
 
Private Sub txtJobNr_KeyDown(KeyCode As Integer, Shift As Integer)
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtJobNr_KeyDown - Begin")
End If
Select Case KeyCode
  Case vbKeyBack          'User back tabbed off of the SLOC field
    If txtJobNr.SelLength = Len(txtJobNr.Text) Then
      ' If SLIC entered and backspace hit, delete everything
      txtJobNr.Text = vbNullString
    End If
 
  Case 9 'tab key
      txtDriverName_GotFocus
    End Select
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtJobNr_KeyDown - End")
End If
End Sub
 
Private Sub txtJobNr_KeyPress(KeyAscii As Integer)
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtJobNr_KeyPress - Begin")
End If
  If (KeyAscii = oKeyDefs.Enter_Key) Then
    EnterAsTab KeyAscii
    KeyAscii = KeyToUpperCase(KeyAscii)
    Exit Sub
  End If
 
  If (KeyAscii = 9 And cboDow.Enabled = False And ctlWkendingDate.Enabled = False) Then
      txtJobNr_Validate (True)
  End If
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtJobNr_KeyPress - End")
End If
End Sub
 
Private Sub txtJobNr_LostFocus()
Dim i As Integer
    On Error GoTo ErrorHandler
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtJobNr_LostFocus - Begin")
End If
 
txtJobNr.BackColor = vbWhite
If IsBlank(txtJobNr.Text) And Me.ActiveControl Is ctlWkendingDate Then
       txtDriverName.Enabled = False
       txtTractorNumber.Enabled = False
       txtEloc.Enabled = False
      
       txtDolly.Item(0).Enabled = False
       txtDolly.Item(1).Enabled = False
       txtDolly.Item(2).Enabled = False
      
       chkDriverNotified.Enabled = False
      
       ctlBayNr(0).TabStop = False
       ctlBayNr(1).TabStop = False
       ctlBayNr(2).TabStop = False
      
       grdBayDetails.Item(0).Enabled = False
       grdBayDetails.Item(1).Enabled = False
       grdBayDetails.Item(2).Enabled = False
      
       cboDow.Enabled = True
       ctlWkendingDate.Enabled = True
       ctlWkendingDate.SetFocus
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtJobNr_LostFocus isblank txtJobNr - End")
        End If
       Exit Sub
    ElseIf Not IsBlank(txtJobNr.Text) And Me.ActiveControl Is txtDriverName Then
       txtDriverName.Enabled = True
       txtDriverName.SetFocus
'   ElseIf Not IsBlank(txtJobNr.Text) And Me.ActiveControl Is ctlBayNr Then
    ElseIf Not IsBlank(txtJobNr.Text) And Me.ActiveControl Is txtJobNr Then
    End If
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtJobNr_LostFocus - End")
End If
    Exit Sub
ErrorHandler:
    glErrNum = 400
    gsError = "txtJobNr_LostFocus"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
    End If
End Sub
 
Private Sub txtJobNr_Validate(Cancel As Boolean)
Dim bResult As Boolean
Dim iRc As Integer
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD txtJobNr_Validate - Begin")
End If
    On Error GoTo ErrorExit
  
    If sPreJobNr = txtJobNr.Text And bRval Then
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD txtJobNr_Validate sPreJob = txtJobNr - End")
        End If
       'Already validated.
       If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtJobNr_Validate - End - Already Validated")
       End If
       Exit Sub
    End If
   
    If IsBlank(txtJobNr.Text) Then
       txtDriverName.Enabled = False
       txtTractorNumber.Enabled = False
       txtEloc.Enabled = False
       txtDolly.Item(0).Enabled = False
       txtDolly.Item(1).Enabled = False
       txtDolly.Item(2).Enabled = False
       chkDriverNotified.Enabled = False
       cboDow.Enabled = True
       cboDow.SetFocus
       ctlWkendingDate.Enabled = True
       ctlWkendingDate.SetFocus
       If g_iDebug = 13 Then
          Call InfoLog("frmPDD txtJobNr_Validate isblank txtJobNr - End")
       End If
       Exit Sub
    Else
       txtDriverName.Enabled = True
      ' txtTractorNumber.Enabled = True
      ' txtEloc.Enabled = True
       grdMultLegView.Enabled = True
    End If
   
    iMyRc = 0
    bESC = False
    bExitScreen = False  'When Job# changed, there will be two chances to close the screen.
 
    If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
    End If
 
    sPrevJobNr = txtJobNr.Text
 
    If Not IsBlank(txtJobNr.Text) Then
        pInst.PDInfo.szPDJobNr = txtJobNr.Text
        bResult = pInst.SearchValidJob()
       
        If Not bResult And (bFromGtF8 Or gbFromArrival Or bArrivalTractorOnly) Then
            txtDriverName.Text = vbNullString
            txtTractorNumber.Text = vbNullString
            If txtJobNr.Enabled = False Then txtJobNr.Enabled = True
            txtJobNr.SetFocus
        ElseIf bResult = False Then
            Call ctlToolbarManager.buttonEnabled(BT_F6_UPD_LEG, False)
        End If
       
        cboDow.Enabled = False
        ctlWkendingDate.Enabled = False
    Else
        If Not IsUserShiftTabbing And Not cboDow.Enabled And Not ctlWkendingDate.Enabled Then
           If txtJobNr.Enabled = False Then txtJobNr.Enabled = True
           txtJobNr.SetFocus
           Cancel = True
        Else 'Shift key pressed
           If ctlWkendingDate.Enabled = False Then ctlWkendingDate.Enabled = True
           ctlWkendingDate.SetFocus
           If g_iDebug = 13 Then
               Call InfoLog("frmPDD txtJobNr_Validate blank txtJobNr - End")
           End If
           Exit Sub
        End If
    End If
    If pInst.bJobNumberFound Then
        ' Now, we have the dvr & Trc from tforemp
        If Not IsBlank(pInst.PDInfo.szPDDvrNa) Then
            If (bFromGtF8 Or pInst.bArriveFlag = True) And Trim$(pInst.PDInfo.szPDDvrNa) <> Trim$(txtDriverName.Text) Then
               pInst.PDInfo.szPDDvrNa = txtDriverName.Text
            Else
               txtDriverName.Text = Trim$(pInst.PDInfo.szPDDvrNa)
            End If
        ElseIf Not IsBlank(txtDriverName.Text) And pInst.bArriveFlag Then
            pInst.PDInfo.szPDDvrNa = txtDriverName.Text
        End If
       
        If Not IsBlank(pInst.PDInfo.szPDTrcNr) Then
            If (bFromGtF8 Or pInst.bArriveFlag = True) And Trim$(pInst.PDInfo.szPDTrcNr) <> Trim$(txtTractorNumber.Text) Then
                pInst.PDInfo.szPDTrcNr = txtTractorNumber.Text
            Else
                txtTractorNumber.Text = pInst.PDInfo.szPDTrcNr
            End If
        ElseIf pInst.bArriveFlag And Not IsBlank(txtTractorNumber.Text) Then
            pInst.PDInfo.szPDTrcNr = txtTractorNumber.Text
        ' TT000394 -
        ElseIf Len(txtTractorNumber.Text) = 0 And Len(pInst.PDInfo.szPDTrcNr) > 0 Then
            txtTractorNumber.Text = pInst.PDInfo.szPDTrcNr
        End If
 
        PopulateMultiLegGrid
    End If
   
    'Check if this job number has been predispatched   'Add by QS 2/10/05
    ' must be PD_SUCCESS or NODATA
    If bResult = True And pInst.bArriveFlag <> True Then 'User clicked Yes  QS 1/6/05
       ctlToolbarManager.buttonEnabled BT_F6_UPD_LEG, True
       txtDriverName.Enabled = True
       txtDriverName.SetFocus
       cboDow.Enabled = False  'for tracker 1404
       ctlWkendingDate.Enabled = False
       'DoEvents - Removed 9-16-14 to resolve flow issues with the screen
      
    ElseIf bResult = False Then
       txtJobNr.SelStart = 0
       txtJobNr.SelLength = Len(txtJobNr.Text)
       Cancel = True
    End If
   
    If Not xaMultiLegArray Is Nothing Then
        If xaMultiLegArray.UpperBound(1) <> -1 Then
            If xaMultiLegArray.Value(xaMultiLegArray.UpperBound(1), 8) = False Then
                sPreJobNr = txtJobNr.Text
            Else
                If xaMultiLegArray.Value(grdMultLegView.Bookmark, 8) = True Then
                    Cancel = True
                Else
                    sPreJobNr = txtJobNr.Text
                End If
            End If
        End If
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtJobNr_Validate - End")
    End If
Exit Sub
 
ErrorExit:
glErrNum = 400
    gsError = "txtJobNr_Validate"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Sub
 
Private Sub txtTractorNumber_GotFocus()
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNumber_GotFocus - Begin")
    End If
    If Not pInst.bArriveFlag And Not bFromGtF8 Then
        bExitScreen = False
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNumber_GotFocus - End")
    End If
End Sub
 
 
Private Sub txtTractorNumber_KeyPress(KeyAscii As Integer)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNumber_KeyPress - Begin")
    End If
  If (KeyAscii = oKeyDefs.Enter_Key) Then
    EnterAsTab KeyAscii
    KeyAscii = KeyToUpperCase(KeyAscii)
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNumber_KeyPress - End - If (KeyAscii = oKeyDefs.Enter_Key)")
    End If
    Exit Sub
  End If
  KeyAscii = KeyToUpperCase(KeyAscii)
  If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNumber_KeyPress - End")
    End If
End Sub
 
Private Sub txtTractorNumber_LostFocus()
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNumber_LostFocus - Begin")
    End If
    bESC = False
    bExitScreen = False
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNumber_LostFocus - End")
    End If
End Sub
 
Private Sub txtTractorNumber_Validate(Cancel As Boolean)
    Dim iResult As Integer
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNr_Validate - Begin")
    End If
   
    On Error GoTo ErrorExit
    If Not txtTractorNumber.Enabled Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD txtTractorNr_Validate tractor number enabled - End")
        End If
       Exit Sub
    End If
 
    If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
    End If
   
    If txtTractorNumber.Text = txtLastTractorNumber And Not IsBlank(txtLastTractorNumber) Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD txtTractorNr_Validate tractorNumber = lastTractorNumber - End")
        End If
        Exit Sub
    End If
   
    If GetKeyState(vbEnter) = 0 Then
        bExitScreen = False
    End If
   
    If (pInst.bArriveFlag Or bFromGtF8) And pInst.PDInfo.hPDHll.Count = 0 And bExitScreen Then
        iResult = pInst.GetSchedJobs()
    End If
   
    pInst.PDInfo.szPDTrcNr = Trim$(txtTractorNumber.Text)
    If Len(pInst.PDInfo.szPDTrcNr) < MIN_TRACTORNR_SIZE Then
        pInst.PDInfo.szPDTrcNr = ""
       
        Cancel = True
        txtLastTractorNumber = ""
        ctlWarningPopup.ShowWarning txtTractorNumber.hwnd, LoadResString(sMinimumlengthoftractornumber), 1800
    Else
        ' Search if Tractor exists on property
        If oClsDb.SelectExistTractor(pInst.PDInfo) Then
            pInst.PDInfo.bTrcVisited = True
        Else
            Cancel = True
        End If
    End If
   
    If Cancel Then
        If txtTractorNumber.Enabled Then
           txtTractorNumber.SetFocus
           txtTractorNumber.SelStart = 0
           txtTractorNumber.SelLength = Len(txtTractorNumber.Text)
        End If
        txtLastTractorNumber = ""
    Else
        txtLastTractorNumber = txtTractorNumber.Text
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD txtTractorNr_Validate - End")
    End If
   
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "txtTractorNumber_Validate"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
End Sub
 
Private Sub tbrStd_ButtonClick(ByVal Button As MSComctlLib.Button)
 
    On Error GoTo ErrorExit
   
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD tbrStd_ButtonClick - Begin")
    End If
   
    With ctlToolbarManager
        Select Case .buttonType(Button)
        Case BT_F1_HELP
            mnuHelp_Click
        Case BT_F2_TLR_LD_MSG
            ShowLdTlrMsg
        Case BT_F3_DRIVER_SCHED
            DoDspSchedByDriver
        Case BT_S_F3_SCHED_BY_LOAD
            DoDspSchedByLoadName
        Case BT_F4_MULTI_BAY
            DoMultiBayDisplay
        Case BT_F5_QUICK_SEARCH
            DoQuickSort
        Case BT_F6_UPD_LEG
            mnuF6AddUpdateLeg_Click
        Case BT_F7_MULTI_LOAD
            DoDspMultiLd
        Case BT_F8_PREDIS
            'Team Track #177 - if the job was predispatched, send both ELOC and Driver Name.
            'Otherwise, send just the driver name
            If g_bDriverPredispatched = False Then
                GetPreDScreen PDDISPLAY_BOTH, txtEloc.Text, txtDriverName.Text, txtJobNr.Text
            Else
                GetPreDScreen CLVT_DRIVER, txtEloc.Text, txtDriverName.Text, txtJobNr.Text
            End If
        Case BT_S_F8_NON_PREDIS
            GetNonPreDScreen txtEloc.Text
        Case BT_F9_DOLLIES_ON_PROP
            GetDollyScreen txtEloc.Text
        Case BT_F10_ACCEPT
            If mnuF10Accept.Enabled = True Then
                DoSaveData
            End If
        Case BT_F11_CANCEL_PD
            mnuF11CancelPredispatch_Click
        Case BT_CTL_P_PRINT
            Print_to_Printer
        Case BT_CTL_X_EXIT
             'Unload Me
            'WR01024 ADID:-dmx7plm - Start
                Call mnuClose_Click
            'WR01024 ADID:-dmx7plm - End
        Case BT_F1_HELP
            Form_Help Me
        End Select
    End With
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD tbrStd_ButtonClick - End")
    End If
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "tbrStd_ButtonClick"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Sub
 
Private Sub InitPdInfo()
Dim sLocHub As String
Dim sRc As Integer
Dim i As Integer
Dim bCancel As Boolean
 
On Error GoTo ErrorHandler
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD InitPdInfo - Begin")
    End If
    Set pInst = New PddApp
    sLocHub = oWSNotify.LocalHub
    pInst.PDInfo.szCurrentCny = Left$(sLocHub, 2)
    pInst.PDInfo.szCurrentSlic = Right$(sLocHub, 5)
    pInst.bArriveFlag = gbFromArrival
   
    If pInst.bArriveFlag Then
        'Enable - so when we go into validate functions, it will validate instead of exiting
        txtJobNr.Enabled = True
        txtDriverName.Enabled = True
        txtTractorNumber.Enabled = True
       
        'TT#3330 - app2dmw - 3-25-2013
        frmBayData.Enabled = True
       
        txtJobNr.Text = raArrivalData(1).sActFdrJobNa
        txtDriverName.Text = raArrivalData(1).sActDvrNa
        txtTractorNumber.Text = raArrivalData(1).sActTrcNr
       
        pInst.PDInfo.szPDJobNr = raArrivalData(1).sActFdrJobNa
        pInst.PDInfo.szPDDvrNa = raArrivalData(1).sActDvrNa
        pInst.PDInfo.szPDJobDomCn = raArrivalData(1).sFdrJobDmcCnyCd
        pInst.PDInfo.szPDJobDomSlic = raArrivalData(1).sFdrJobDmcSAB
               
 
        szFdrSchWndDT = Format(ctlWkendingDate.Value, "MM/DD/YYYY")
        For i = 0 To 6
            If g_sWeekCode(i) = Trim$(cboDow.Text) Then
                Exit For
            End If
        Next i
        szFdrSchDow = i
        bCancel = False
        'this will fill in any structures that is needed, TSAT #4662
        Call txtJobNr_Validate(bCancel)
        If Not (bCancel And (bFromGtF8 Or gbFromArrival Or bArrivalTractorOnly)) Then
            Call txtDriverName_Validate(bCancel)
            Call txtTractorNumber_Validate(bCancel)
        End If
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD InitPdInfo - End")
    End If
    Exit Sub
ErrorHandler:
    glErrNum = 400
    gsError = "InitPdInfo"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
    End If
 
End Sub
 
Private Sub SetScreenDetails()
Dim i As Integer, j As Integer
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD SetScreenDetails - Begin")
    End If
    On Error GoTo ErrorExit
 
    'dsb ~~~ The next statement must be changed from PreDispatchedLoads to
    '                                                PredispatchDriver
    If gbReadOnlyMode Then
        Me.Caption = OGlobalFormNames.PredipatchDriver & LoadResString(125)
    Else
        Me.Caption = OGlobalFormNames.PredipatchDriver
    End If
    tbrStd.Enabled = True
   
    ' Set up menu caption structures
    mnuFile.Caption = LoadResString(S1031)
    mnuPrint.Caption = LoadResString(S10311)
    mnuPrint_To_Printer.Caption = LoadResString(S103111)
    mnuPrint_To_Spreadsheet.Caption = LoadResString(S103112)
    mnuPrint_To_File.Caption = LoadResString(S103113)
    mnuReset.Caption = "Reset Form" & vbTab & "Esc"
    mnuClose.Caption = LoadResString(S10312)
    mnuFunctionKeys.Caption = LoadResString(S1033)
    mnuF2TrailerMessage.Caption = LoadResString(S10331)
    mnuF3JobSchedules.Caption = LoadResString(S10332)
    mnuSchByLds.Caption = LoadResString(S10333)
    mnuF4MultiBayDisplay.Caption = LoadResString(S10334)
    mnuF5QuickSort.Caption = LoadResString(S10335)
    mnuF7MultipleLoad.Caption = LoadResString(S10336)
    mnuPDLdsDisplay.Caption = LoadResString(S10337)
    mnuNPDLdsDisplay.Caption = LoadResString(S10338)
    mnuDolliesDisplay.Caption = LoadResString(S10339)
    mnuF10Accept.Caption = LoadResString(S1033A) & vbTab & "F10"
    mnuView.Caption = LoadResString(S1034)
    mnuToolbar.Caption = LoadResString(S10341)
    mnuStandardToolbar.Caption = LoadResString(S103411)
    mnuHelp.Caption = LoadResString(S1035)
   
    mnuStandardToolbar.Checked = True
    mnuPrint.Enabled = True
 
    'Setup Labels
   lblJobNumber.Caption = LoadResString(sRES_JOB_NR)
    lblDriverName.Caption = LoadResString(sRES_DRIVER_NAME)
    lblTractorNumber.Caption = LoadResString(sRES_TRACTOR)
    lblEloc.Caption = LoadResString(sRES_ELOC)
    lblDriverNotified.Caption = LoadResString(sRES_DRIVER_NOTIFY)
   
    txtJobNr.Text = ""
    txtDriverName.Text = ""
    txtTractorNumber.Text = ""
    txtEloc.Text = ""
    Set oGridFunctions = New HFCSGlobalUtils.clsGridFunctions
 
    With oGridFunctions
        Call .GridInitialize(Me.grdMultLegView, GridType.ReadOnly)
        Call .Set_Grid_Properties(Me.grdMultLegView, GridType.ReadOnly)
       
     '   Call .GridAddSplit(Me.grdMultLegView, 1, "Load", dbgNone, dbgNumberOfColumns, 6, dbgCenter)
       
        Call .GridAddSplit(Me.grdMultLegView, _
                            0, _
                            "Select Leg to Predispatch", _
                            dbgAutomatic, _
                            dbgScalable, _
                            10, dbgCenter)
                           
        Call .GridAddColumn(Me.grdMultLegView, 0, "Leg Sched Dep Time", 0, 1675, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 1, "ELOC", 0, 1005, "", True, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 2, "LEGOID", 0, 1, "", False, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 3, "DriverName", 0, 1, "", False, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 4, "Domicile", 0, 1, "", False, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 5, "DomicileCny", 0, 1, "", False, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 6, "TractorNumber", 0, 1, "", False, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 7, "ELOCCny", 0, 1, "", False, False, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 8, "Departed", 0, 1, "", False, True, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        Call .GridAddColumn(Me.grdMultLegView, 9, "JOBOID", 0, 1, "", False, True, True, False, dbgCenter, dbgCenter, ReadOnly, True)
        grdMultLegView.FetchRowStyle = True
    End With
   
    For i = 0 To 2
        ' Define CtlBay Control
       
        ctlBayNr(i).Top = grdBayDetails(i).Top + grdBayDetails(i).RowTop(0)
        ctlBayNr(i).Height = grdBayDetails(i).RowHeight
        ctlBayNr(i).Left = grdBayDetails(i).Left + grdBayDetails(i).Splits(1).RecordSelectorWidth + grdBayDetails(i).Splits(1).Columns(GRD1_COL_BAY).Left + Screen.TwipsPerPixelX
        ctlBayNr(i).Width = grdBayDetails(i).Splits(0).Columns(GRD1_COL_BAY).Width + Screen.TwipsPerPixelX
        ctlBayNr(i).ZOrder 1
       
        ctlBayNr(i).Appearance = 0
 
       
'~~~###~~~ enable once Bay control has changed!!!
       ctlBayNr(i).GateBay = gbFromArrival
       
        ' Define grdBayDetails
        With grdBayDetails(i).Columns
            .Item(GRD1_COL_BAY).Caption = LoadResString(sRES_BAY_NUM)
            .Item(GRD1_COL_POS).Caption = LoadResString(sRES_POS)
            .Item(GRD1_COL_TRAILER).Caption = LoadResString(sRES_TRAILER_NUM)
            .Item(GRD1_COL_TYPE).Caption = LoadResString(sRES_TRAILER_TYPE)
            .Item(GRD1_COL_ORIG).Caption = LoadResString(sRES_ORIG)
            .Item(GRD1_COL_OS).Caption = LoadResString(sRES_ORIG_SORT)
            .Item(GRD1_COL_DEST).Caption = LoadResString(sRES_DEST)
            .Item(GRD1_COL_DS).Caption = LoadResString(sRES_DEST_SRT)
            .Item(GRD1_COL_SEQ).Caption = LoadResString(sRES_LD_SEQ)
            .Item(GRD1_COL_LDCD).Caption = LoadResString(sRES_LD_CD)
            .Item(GRD1_COL_RI).Caption = LoadResString(sRES_RT)
            .Item(GRD1_COL_PCS).Caption = LoadResString(sRES_PCS)
            .Item(GRD1_COL_PER).Caption = LoadResString(sRES_PCT)
            .Item(GRD1_COL_TAG).Caption = LoadResString(sRES_TAG)
            .Item(GRD1_COL_CREATE).Caption = LoadResString(sRES_CREATE)
            .Item(GRD1_COL_DUE).Caption = LoadResString(sRES_DUE_DATE)
            .Item(GRD1_COL_DEP_TM).Caption = LoadResString(sRES_DEP_TIME)
           ' .Item(GRD1_COL_HAZMAT).Caption = LoadResString(sRES_HAZMAT)
            .Item(GRD1_COL_REMARKS).Caption = LoadResString(sRES_REMARKS)
           
            For j = 0 To .Count - 2
                .Item(i).Alignment = dbgCenter
            Next j
           
        End With
        With grdBayDetails(i).Splits
            .Item(GRD1_SPL_TRAILER).Caption = LoadResString(sRES_TRAILER)
            .Item(GRD1_SPL_LOAD).Caption = LoadResString(sRES_LOAD_NAME)
            .Item(GRD1_SPL_VOLUME).Caption = LoadResString(sRES_VOL)
            .Item(GRD1_SPL_DATE).Caption = LoadResString(sRES_DATE)
            For j = 0 To .Count - 1
                .Item(j).MarqueeStyle = dbgNoMarquee
            Next j
        End With
                
        ' Set Dolly constants
        lblDolly(i).Caption = LoadResString(sRES_DOLLY)
        txtDolly(i).Text = ""
 
    Next
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD SetScreenDetails - End")
    End If
   
    Set oGridFunctions = Nothing
Exit Sub
ErrorExit:
    glErrNum = 400
    gsError = "SetScreenDetails"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Sub
 
Public Sub DisplayLoadInfo(pData As PdStruct, LdIndex As Integer)
    On Error GoTo ErrorHandler
' DisplayLoadInfo: This function will take the information found within
'                   the PDSTRUCT Record and display it on the screen.  If
'                   the Country Code matches the Current HUB COUNTRY CODE, don't
'                   display it.
'
'                   LdIndex is zero based.
'
Dim szDolNr As String    ' dolly number
Dim i As Integer
Dim iBay As Integer
Dim lFOPDollyKey As Long ' Dolly entity Key
Dim oColBays As HFCSYardObject.Bay
   
    
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD DisplayLoadInfo - Begin")
    End If
    Set oColBays = New HFCSYardObject.Bay
   
    i = LdIndex
      
    With pData
   
        ctlBayNr(i).Text = .szBayNr
       
        If Len(.szBayNr) > 0 And IsNumeric(.szBayNr) Then
            LockBayForPredispatch (i)
        End If
       
        If IsNull(pData.ExtraDolly) And i < 2 Then
            chkExtraDolly(i).Value = 0
        ElseIf i < 2 Then
            If pData.ExtraDolly = "0" Then
                chkExtraDolly(i).Value = 0
            ElseIf pData.ExtraDolly = "1" Then
                chkExtraDolly(i).Value = 1
            End If
        End If
       
        If (i = 0) Then
            .SameLoadName = False
        End If
       
        Set gpData = pData
        ' Force the grid to display
        grdBayDetails(i).ReBind
 
    End With
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD DisplayLoadInfo - End")
    End If
    Exit Sub
ErrorHandler:
    glErrNum = 400
    gsError = "DisplayLoadInfo"
   
    'the object has disconnect from its clients grid error, retry the grid rebind
    If Err.Number = -2147417848 Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD DisplayLoadInfo - received error -2147417848")
        End If
       
        If RecoverGridError(grdBayDetails(i)) = True Then
            Resume Next
        End If
    End If
   
    Call oProc.update_error_object(Me, gsError)
   
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        Set oColBays = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
 
End Sub
 
Public Sub ClearLoadDisplay(LdIndex As Integer)
 
    On Error GoTo ErrorHandler
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearLoadDisplay - Begin - LdIndex: " & LdIndex)
    End If
    On Error Resume Next
    Set gpData = Nothing
   
    If pInst.PDInfo.hPDHll.Count > LdIndex + 1 And iMyRc <> 0 And LdIndex = 2 Then  'QS 8/1
        pInst.PDInfo.hPDHll.Remove LdIndex + 1
    End If
   
    'TT00537 - when this index is greater than 2, causes screen to crash, just exit
    If LdIndex > 2 Then Exit Sub
   
    grdBayDetails(LdIndex).ReBind
   
'    If LdIndex > 0 Then
'        bDataFromEloc(LdIndex - 1) = bDataFromEloc(LdIndex)
'    End If
   
    ' clear the attached dolly info
    txtDolly(LdIndex).ForeColor = vbWindowText
    txtDolly(LdIndex).Text = ""
    txtDolly(LdIndex).Tag = 0
    ' TT000394 - There are only two extra dolly controls.
    If LdIndex < 2 Then
        chkExtraDolly(LdIndex).Value = 0
        chkExtraDolly(LdIndex).Enabled = False
        chkExtraDolly(LdIndex).TabStop = False
    End If
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearLoadDisplay - End - LdIndex: " & LdIndex)
    End If
Exit Sub
ErrorHandler:
    'the object has disconnect from its clients grid error, retry the grid rebind
    If Err.Number = -2147417848 Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearLoadDisplay - received error -2147417848")
        End If
       
        If RecoverGridError(grdBayDetails(LdIndex)) = True Then
            Resume Next
        End If
    End If
End Sub
 
Public Function GetAttatchedDolly(LdIndex As Integer, _
                                      lFopTlrEntity As Long, _
                                      szBayNr As String, _
                                ByRef szDolNr As String, _
                                ByRef lDollyKey As Long) As Boolean
On Error GoTo ErrorHandler
Dim bFound As Boolean
Dim i As Integer
Dim colDol As Collection
Dim colTlr As Collection
Dim oDol As HFCSDollyObject.Dolly
Dim oTlr As HFCSTrailerObject.Trailer
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetAttachedDolly - Begin - LdIndex: " & LdIndex)
    End If
    szDolNr = ""
    lDollyKey = 0
    ' 1. call 9101 for bay#
   
    Set oBay = oYard.BayGetSingle(szBayNr, HFCS_INDICATOR_FALSE, 13, oClsDb.DBClass)
    If oBay Is Nothing Then
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD GetAttachedDolly oBay nothing - End - LdIndex: " & LdIndex)
        End If
        GetAttatchedDolly = False
        Exit Function
    End If
    Set colTlr = oBay.Equipment(BAY_EQ_TRAILER)
    Set colDol = oBay.Equipment(BAY_EQ_DOLLY)
   
    ' 2. look for a foptlr with the fop-ent-key
    ' Go through the nodes in the HLL 1 by 1 and check for
    ' compatible RECTYPEs.
    bFound = False
    i = 1
    If Not colTlr Is Nothing Then
        Do While Not bFound And i <= colTlr.Count
            Set oTlr = colTlr.Item(i)
            ' Is this the node we are looking for ??
            If oTlr.FOPEntityKey = lFopTlrEntity Then
                ' 3. get linkageKey & TlrPos from fop-tlr-node
                bFound = True
            End If
            i = i + 1
        Loop
    End If
   
    If bFound Then
        ' 4. look for fopdol node with same pos-cd & linkage-key
        i = 1
        ' Go through the nodes in the HLL 1 by 1 and check for
        ' compatible RECTYPEs.
        bFound = False
        If Not colDol Is Nothing Then
            Do While Not bFound And i <= colDol.Count
                Set oDol = colDol.Item(i)
                ' Is this the node we are looking for ??
                If oDol.BayNumber = oTlr.BayNumber _
                        And oDol.DollyAttachmentPoint = oTlr.TrailerPosition Then
                    bFound = True
                    szDolNr = oDol.DollyNumber
                    lDollyKey = oDol.FOPEntityKey
                End If
                i = i + 1
            Loop
        End If
    End If
   
    
    Set colTlr = Nothing
    Set colDol = Nothing
   
    ' Return true if found
    GetAttatchedDolly = bFound
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetAttachedDolly  - End - LdIndex: " & LdIndex)
    End If
    Exit Function
ErrorHandler:
    glErrNum = 400
    gsError = "GetAttatchedDolly"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        GetAttatchedDolly = False
        Set oEventlog = Nothing
        Set colDol = Nothing
        Set colTlr = Nothing
        Set oDol = Nothing
        Set oTlr = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Function
 
Public Sub DisplayLdData(hllLds As hPDHll)
    On Error GoTo ErrorHandler
    Dim i As Integer
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DisplayLdData - Begin")
    End If
 
    'Display data
    For i = 1 To hllLds.Count
        DisplayLoadInfo hllLds.Item(i), i - 1
    Next i
   
    'Clear grid
    If hllLds.Count = 0 Then
        For i = 1 To 3
'           chkCancelPD(i - 1).Enabled = False
'           chkCancelPD(i - 1).Value = 0
           ClearLoadDisplay i - 1
        Next i
       Exit Sub
    End If
    'Clear data that we don't need any more from the screen
    If iBayIndex + 1 < 3 Then
      If (iMyRc = vbKeyDelete Or iMyRc = UMBID_DELETE) And Len((ctlBayNr.Item(iBayIndex + 1).Text)) > 0 Then
         For i = hllLds.Count + 1 To 3
            ClearLoadDisplay i - 1
'            chkCancelPD(i - 1).Enabled = False
'            chkCancelPD(i - 1).Value = 0
            iMyRc = 0
         Next i
      Else
         For i = hllLds.Count + 1 To hllLds.Count + 1 Step -1
            If i > 1 Then
               ClearLoadDisplay i - 1
'               chkCancelPD(i - 1).Enabled = False
'               chkCancelPD(i - 1).Value = 0
            Else
               grdBayDetails(iBayIndex).ReBind
            End If
         Next i
      End If
    Else
       For i = hllLds.Count + 1 To iBayIndex + 1
          ClearLoadDisplay i - 1
'          chkCancelPD(i - 1).Enabled = False
'          chkCancelPD(i - 1).Value = 0
       Next i
    End If
    sServiceTypeCD = ""
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD DisplayLdData - End")
        End If
    Exit Sub
   
ErrorHandler:
    glErrNum = 400
    gsError = "DisplayLdData"
    'the object has disconnect from its clients grid error, retry the grid rebind
    If Err.Number = -2147417848 Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD DisplayLdData - received error -2147417848")
        End If
       
        If RecoverGridError(grdBayDetails(iBayIndex)) = True Then
            Resume Next
        End If
    End If
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
    End If
 
End Sub
 
Public Sub ExitDlg()
    If g_iDebug = 13 Then
           Call InfoLog("frmPDD ExitDlg - Begin")
        End If
    Unload Me
    If g_iDebug = 13 Then
           Call InfoLog("frmPDD ExitDlg - End")
    End If
End Sub
 
Public Sub ClearPddData()
Dim i As Integer
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearPddData - Begin")
        End If
   
    For i = 1 To 3
        ClearLoadDisplay i - 1
    Next
 
    txtDriverName.Text = ""
    txtEloc.Text = ""
    txtJobNr.Text = ""
    txtTractorNumber.Text = ""
    g_bShowMatchingLoad = True
    Set pInst = Nothing
    InitPdInfo
    mbElocValid = False
    mbClearData = False
   
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearPddData - End")
        End If
 
End Sub
 
Public Sub ProcessRetData(uHFCSLoad As GLOBALDEFS.HFCSLOAD)
Dim sSab As String, sCny As String
Dim sNoEquip As String
Dim pData As PdStruct
Dim bResult As Boolean
Dim iRc As Integer
Dim lResult As Long
Dim i As Integer
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ProcessRetData  - Begin")
    End If
   
    If IsBlank(uHFCSLoad.BayNumber) Then
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD ProcessRetDAta blank uHFCSLoad.BayNumber - End")
        End If
        Exit Sub
   
    End If
   
    ' If an empty bay was selected, we have nothing to do
    If IsBlank(uHFCSLoad.TrailerNumber) Then
        ' Bay is empty : show popup #73
        sNoEquip = LoadResString(sBay) & uHFCSLoad.BayNumber & LoadResString(sIsempty)
        ctlWarningPopup.ShowWarning ctlBayNr.Item(pInst.icdl - 1).hwnd, sNoEquip, 2000
        bReLoad = False
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD ProcessRetDAta blank uHFCSLoad.TrailerNumber - End")
        End If
        Exit Sub
    End If
   
    Set pData = New PdStruct
   
    ' look up all of the bay/load data based on the bay/trailer/load data
    With uHFCSLoad
        lResult = oClsDb.RetrieveDataByBayLd(.BayNumber, .TrailerNumber, .OriginSlic, _
                                    .OriginSort, .DestinationSlic, .DestinationSort, _
                                    .SequenceNumber, _
                                    pData)
    End With
   
    If lResult <> DB_SUCCESS Then
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD ProcessRetDAta lResult - End")
       End If
        Exit Sub
    End If
   
 
    If Not gbProcessRetData Then
       If pInst.PDInfo.hPDHll.Count = 0 Then
          pInst.icdl = pInst.PDInfo.hPDHll.Count + 1
       Else
          pInst.icdl = pInst.PDInfo.hPDHll.Count
       End If
    End If
    pInst.icdl = iBayIndex + 1
    pInst.PDInfo.hPDHll.OverLayBayLdData pData, pInst.icdl
    ' This is correct
    bResult = pInst.ProcessBaySched(pInst.icdl)  'QS 061305
   
    If bResult Then
        'QS test 5/9/05
        For i = 1 To pInst.PDInfo.hPDHll.Count
            If pInst.PDInfo.hPDHll.Count >= iBayIndex + 1 Then   '8/12/05
                DisplayLoadInfo pInst.PDInfo.hPDHll.Item(iBayIndex + 1), iBayIndex
            End If
        Next i
        If pData.szMultLdIr = True And (pData.lTlrMsgId = 0 And pData.lLdMsgId = 0) Or _
        ((pData.lTlrMsgId > 0 Or pData.lLdMsgId > 0) And pData.szMultLdIr = 1) Then 'QS make f7 button enable  12/08/04
           SetAvailMenuAndTool mat_bay, True, True
           m_bEditTrailerMessage = True
           m_bEditLoadMessage = True
        ElseIf pData.szMultLdIr = 0 And (pData.lTlrMsgId > 0 Or pData.lLdMsgId > 0) Then
           SetAvailMenuAndTool mat_bay, True, False
           m_bEditLoadMessage = True
           m_bEditTrailerMessage = True
        ElseIf Len(pData.szBayNr) > 0 Then
           SetAvailMenuAndTool mat_bay, True, False
           m_bEditLoadMessage = True
           m_bEditTrailerMessage = True
        Else
           SetAvailMenuAndTool mat_bay, False, False
        End If
       
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ProcessRetDAta - End")
    End If
    Exit Sub
   
Error_Handler:
  glErrNum = 400
  gsError = "ProcessRetData"
  Call oProc.update_error_object(Me, gsError)
 
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
End If
End Sub
 
Public Function PopulateLoad(PdData As PdStruct) As Variant
'=========================================================
' Sub:          populateLoad()
' Description:  for calling trailer/load messages
'=========================================================
Dim curLoad As GLOBALDEFS.HFCSLOAD
   
    On Error GoTo populateLoadError
    'oProc.update_error_object Me, "PopulateLoad"
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD PopulateLoad - Begin")
    End If
    lTlrMsgId = 0
    lLdMsgId = 0
   
    curLoad.BayNumber = PdData.szBayNr
    curLoad.TrailerNumber = PdData.szTlrNr
    curLoad.OriginSlic = PdData.szOrigin
    curLoad.OriginSort = PdData.szOriginSrt
    curLoad.DestinationSlic = PdData.szDestin
    curLoad.DestinationSort = PdData.szDestinSrt
    If Len(CStr(PdData.iSequenceNr)) = 1 Then
       curLoad.SequenceNumber = "0" & PdData.iSequenceNr
    Else
       curLoad.SequenceNumber = PdData.iSequenceNr
    End If
    curLoad.TrailerMessageNumber = PdData.lTlrMsgId
    curLoad.LoadMessageNumber = PdData.lLdMsgId
   
    If IsBlank(sLoadMsg) Then
       curLoad.LoadMessage = CStr(oClsDb.SelectMsg(PdData.lLdMsgId))
    Else
       curLoad.LoadMessage = sLoadMsg
    End If
    If IsBlank(sTrailerMsg) Then
       curLoad.TrailerMessage = CStr(oClsDb.SelectMsg(PdData.lTlrMsgId))
    Else
       curLoad.TrailerMessage = sTrailerMsg
    End If
    PopulateLoad = curLoad
   
    lTlrMsgId = PdData.lTlrMsgId
    lLdMsgId = PdData.lLdMsgId
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD PopulateLoad - End")
    End If
 
Exit Function
 
populateLoadError:
    glErrNum = 400
    gsError = "PopulateLoad"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Function
 
Private Sub AcceptPreDispatch()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD AcceptPreDispatch - Begin")
    End If
    MsgBox LoadResString(sSavingData)
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD AcceptPreDispatch - End")
    End If
End Sub
 
Private Sub SetupToolbar()
   
    On Error GoTo ErrorExit
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetupToolbar - Begin")
    End If
       
    Set ctlToolbarManager.Toolbar = tbrStd
   
    With ctlToolbarManager
        .addButton BT_F2_TLR_LD_MSG
        .addButton BT_F3_DRIVER_SCHED
        .addButton BT_S_F3_SCHED_BY_LOAD
        .addButton BT_F4_MULTI_BAY
        .addButton BT_F5_QUICK_SEARCH
        .addButton BT_F6_UPD_LEG
        .addButton BT_F7_MULTI_LOAD
        .addButton BT_F8_PREDIS
        .addButton BT_S_F8_NON_PREDIS
        .addButton BT_F9_DOLLIES_ON_PROP
        .addButton BT_F10_ACCEPT
        .addButton BT_CTL_P_PRINT
        .addButton BT_CTL_X_EXIT
        .addButton BT_F1_HELP
        .addButton BT_F11_CANCEL_PD
       
        .buttonEnabled BT_F2_TLR_LD_MSG, False
        .buttonEnabled BT_F3_DRIVER_SCHED, False
        .buttonEnabled BT_S_F3_SCHED_BY_LOAD, False
 '<WR1024><ADID: dmx7plm> - Start
        .buttonEnabled BT_F4_MULTI_BAY, True
        .buttonEnabled BT_F5_QUICK_SEARCH, True
'<WR1024><ADID: dmx7plm> - End
 
        .buttonEnabled BT_F6_UPD_LEG, False
        .buttonEnabled BT_F7_MULTI_LOAD, False
        .buttonEnabled BT_F8_PREDIS, False
        .buttonEnabled BT_S_F8_NON_PREDIS, False
        .buttonEnabled BT_F9_DOLLIES_ON_PROP, False
        .buttonEnabled BT_F10_ACCEPT, False
        .buttonEnabled BT_F11_CANCEL_PD, False
    End With
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetupToolbar - End")
    End If
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "SetupToolbar"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
   '     Unload Me
    End If
End Sub
 
Public Function LockBayForPredispatch(Index As Integer) As Boolean
On Error GoTo ErrorHandler
Dim i As Integer
Dim szBayNr As String
Dim iBayNr As Integer
'Dim oBayInfo   As HFCSYardObject.Bay
Dim oBay2  As HFCSYardObject.Bay
Dim bAddCollection As Boolean
Dim rsBays As ADODB.Recordset
Dim sMsg As String
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD LockBayForPredispatch - Begin - Index: " & Index)
    End If
 
    bAddCollection = False
    szBayNr = ctlBayNr(Index).Text
    iBayNr = Int(Val(szBayNr))
   
    ' Check if iBayNr is valid
    Set oBay = GetBayInformation(szBayNr)
   
    If oBay Is Nothing Then
        LockBayForPredispatch = False
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD LockBayForPredispatch oBay is nothing - End - Index: " & Index)
        End If
        Exit Function
    End If
   
    ' Check if bay is marked unavailable
    If oBay.Available = HFCS_INDICATOR_FALSE Then
        sMsg = LoadResString(sBay) & szBayNr & LoadResString(sIsnotavailableatthistime)
        ctlWarningPopup.ShowWarning ctlBayNr(Index).hwnd, sMsg, 2000
        LockBayForPredispatch = False
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD LockBayForPredispatch obay.available = false - End - Index: " & Index)
        End If
        Exit Function
    End If
   
    ' Lock the new bay, if bay is already locked by somebody else
    If ProcessLockBay(szBayNr, oBay, BAY_LOCK) <> SUCCESSFUL_LOCK Then
        sMsg = LoadResString(sBay) & szBayNr & LoadResString(sIsnotavailableatthistime)
        ctlWarningPopup.ShowWarning ctlBayNr(Index).hwnd, sMsg, 2000
        LockBayForPredispatch = False
        ctlBayNr(Index).Text = szBayNr
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD LockBayForPredispatch - processlockbay failed - End - Index: " & Index)
        End If
        Exit Function
    End If
   
    pInst.raLockedBay(Index + 1) = iBayNr
   
    'Add the bay info in a collection
    If g_colBays.Count > 0 Then
       For Each oBay2 In g_colBays
          If oBay2.Number = oBay.Number Then
             bAddCollection = False
             Exit For
          Else
             Set rsBays = oClsDb.CheckForLockedBay(oBay2.Number)
            
             'this has to be done because the Yard Object unlocks the bay when setting the bay object to another bay
             'therefore, lock other bays that show if it is not locked.
             If IsNull(rsBays.Fields("bay_lck_ts").Value) Then
                'check to see if bay is locked
                oBay2.LockBay BAY_PRIMARY_LOCK
             End If
             bAddCollection = True
          End If
       Next
       If bAddCollection Then
          Call g_colBays.Add(oBay, oBay.Number)
       End If
    Else
       Call g_colBays.Add(oBay, oBay.Number)
    End If
   
    LockBayForPredispatch = True
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD LockBayForPredispatch - End - Index: " & Index)
    End If
 
    Exit Function
ErrorHandler:
    glErrNum = 400
    gsError = "LockBayForPredispatch"
    Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Call oClsDb.DBClass.CloseRecordSet(rsBays)
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
    End If
 
End Function
 
Private Function GetBayInformation(BayNr As String) As HFCSYardObject.Bay
' Retrieve information regarding a specified bay.
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetBayInformation - Begin - BayNr: " & BayNr)
    End If
  '  Set oBay = Nothing
    Set GetBayInformation = oYard.BayGetSingle(BayNr, HFCS_INDICATOR_FALSE, 13, oClsDb.DBClass)
  '  Set GetBayInformation = oBay
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetBayInformation - End - BayNr: " & BayNr)
    End If
End Function
 
Private Function ProcessLockBay(szBayNr As String, _
                              ByRef objBay As HFCSYardObject.Bay, _
                              requestedBayStatus As Integer _
                              ) As Integer
 
On Error GoTo ErrorHandler
 
Dim iRc As Integer
Dim cnt As Long
Dim szMsg As String
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ProcessLockBay - Begin")
    End If
 
cnt = 0
 
    If objBay Is Nothing Then
        If IsBlank(szBayNr) Then
            ProcessLockBay = LOCK_ERROR
            Exit Function
        End If
       
        Set objBay = oYard.BayGetSingle(szBayNr, HFCS_INDICATOR_FALSE, 13, oClsDb.DBClass)
        If objBay Is Nothing Then
            ProcessLockBay = LOCK_ERROR
            Exit Function
        End If
    End If
    Select Case requestedBayStatus
    Case BAY_LOCK
        ProcessLockBay = IIf(objBay.LockBay(BAY_PRIMARY_LOCK), SUCCESSFUL_LOCK, LOCK_ERROR)
    Case BAY_UNLOCK
        objBay.UnlockBay
        ProcessLockBay = SUCCESSFUL_UNLOCK
    Case Else
        ProcessLockBay = LOCK_ERROR
    End Select
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ProcessLockBay - End.")
    End If
   
Exit Function
 
ErrorHandler:
    glErrNum = 400
    gsError = "ProcessLockBay"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
End Function
 
Public Sub ReturnData(FromApp As String, vDataArray() As Variant)
Dim uHFCSLoad As GLOBALDEFS.HFCSLOAD
Dim i As Integer, lb As Integer
Dim sMessage As String
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ReturnData - Begin")
    End If
   
    If bBayCompeleted(iBayIndex) Then
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD ReturnData - End - If bBayCompeleted(iBayIndex): iBayIndex = " & iBayIndex)
        End If
        Exit Sub
    End If
 
    'Reinitail the varibles   QS 12/14/04
    m_bEditTrailerMessage = False
    m_bEditLoadMessage = False
   
    gbProcessRetData = True
    lb = LBound(vDataArray)
   
    
    If FromApp = OGlobalFormNames.NonPreDispatchedLoads Then
        uHFCSLoad = vDataArray(lb)
        ProcessRetData uHFCSLoad
    ElseIf FromApp = OGlobalFormNames.PreDispatchedLoads Then
        uHFCSLoad = vDataArray(lb)
        ProcessRetData uHFCSLoad
    ElseIf FromApp = OGlobalFormNames.Scheduled_Outbounds Then
        uHFCSLoad = vDataArray(lb)
        ProcessRetData uHFCSLoad
    ElseIf FromApp = OGlobalFormNames.ScheduledInformation Then
        uHFCSLoad = vDataArray(lb)
       
        If IsBlank(uHFCSLoad.BayNumber) Then
            sMessage = LoadResString(sTheselectedscheduledloadisnotassoicatedwithabay)
            ctlWarningPopup.ShowWarning ctlBayNr.Item(pInst.icdl - 1).hwnd, sMessage, 2000
            Exit Sub
        End If
   
        ' If an empty bay was selected, we have nothing to do
        If IsBlank(uHFCSLoad.TrailerNumber) Or uHFCSLoad.TrailerNumber = "EMPTY" Then
            ' Bay is empty : show popup #73
            sMessage = LoadResString(sBay) & uHFCSLoad.BayNumber & LoadResString(sIsempty)
           
            ctlWarningPopup.ShowWarning ctlBayNr.Item(pInst.icdl - 1).hwnd, sMessage, 2000 '
            bReLoad = False
            Exit Sub
        End If
       
        ProcessRetData uHFCSLoad
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ReturnData - End.")
    End If
    Exit Sub
   
Error_Handler:
  glErrNum = 400
  gsError = "ReturnData"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Sub
 
Public Sub SetAvailMenuAndTool(Opt As MENU_AND_TOOL_OPTIONS, Optional bShowF2 As Boolean, Optional bShowF7 As Boolean)
Dim bAllowSave As Boolean, i As Integer
 
    On Error GoTo ErrorExit
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetAvailMenuAndTool & Opt = " & Opt & " Begin")
    End If
 
    If Opt = mat_dolly Then
        gbBayReached = True
    End If
 
    'Set Defaults
    mnuDolliesDisplay.Enabled = False
    If gbBayReached Then
        bAllowSave = True
        With pInst
            For i = 1 To .PDInfo.hPDHll.Count
                If Not .bPreDispatched(i) Then
                    bAllowSave = False
                    Exit For
                End If
            Next
        End With
    Else
        bAllowSave = False
    End If
    mnuF10Accept.Enabled = bAllowSave And Not gbReadOnlyMode
    mnuF2TrailerMessage.Enabled = False
    mnuF3JobSchedules.Enabled = False
    mnuF4MultiBayDisplay.Enabled = False
    mnuF5QuickSort.Enabled = False
    mnuF7MultipleLoad.Enabled = False
    mnuNPDLdsDisplay.Enabled = False
    mnuPDLdsDisplay.Enabled = False
    mnuSchByLds.Enabled = False
    mnuF11CancelPredispatch.Enabled = False
   
    mnuPrint.Enabled = True
    
                                               
    Select Case Opt
    Case MAT_BASE_DATA0
        ' All are disabled
        mnuF10Accept.Enabled = False
        mnuF2TrailerMessage.Enabled = bShowF2  ' If the Trailer/load has a msg
        mnuF3JobSchedules.Enabled = True
        mnuF4MultiBayDisplay.Enabled = False
        mnuF5QuickSort.Enabled = False
        mnuF7MultipleLoad.Enabled = bShowF7 ' If the current bay has multi-load
        mnuNPDLdsDisplay.Enabled = False
        mnuPDLdsDisplay.Enabled = False
        mnuSchByLds.Enabled = False
        mnuF11CancelPredispatch.Enabled = False
        mnuPrint.Enabled = False
    Case mat_base_data
        
        mnuF2TrailerMessage.Enabled = bShowF2  ' If the Trailer/load has a msg
        mnuF3JobSchedules.Enabled = True
        mnuF4MultiBayDisplay.Enabled = True
        mnuF5QuickSort.Enabled = True
        mnuF7MultipleLoad.Enabled = bShowF7 ' If the current bay has multi-load
        mnuNPDLdsDisplay.Enabled = True
        mnuPDLdsDisplay.Enabled = True
        mnuSchByLds.Enabled = True
        mnuPrint.Enabled = False
    Case mat_bay
        mnuF2TrailerMessage.Enabled = bShowF2  ' If the Trailer/load has a msg
        mnuF3JobSchedules.Enabled = True
        mnuF4MultiBayDisplay.Enabled = True
        mnuF5QuickSort.Enabled = True
        mnuF7MultipleLoad.Enabled = bShowF7 ' If the current bay has multi-load
        mnuNPDLdsDisplay.Enabled = True
        mnuPDLdsDisplay.Enabled = True
        mnuSchByLds.Enabled = True
          
        mnuDolliesDisplay.Enabled = False   'True
 
        'kg disable - F10 button in the following cases below
        If bDuplicateLoad Then
           mnuF10Accept.Enabled = False
           bDisableF10 = True
        End If
   
        'kg - If any bay or dolly is not valid, disable F10 button
        For i = 0 To 2
           If (Not bBayCompeleted(i) And Len(ctlBayNr.Item(i).Text) > 0) Or (Not bDollyValidated(i) And Len(txtDolly.Item(i).Text) > 0) Then
              If (Not bDollyValidated(i) And Len(txtDolly.Item(i).Text) > 0) Then
                mnuF10Accept.Enabled = False
                bDisableF10 = True
              End If
           ElseIf bBayCompeleted(0) And (Not bBayCompeleted(i) And Len(ctlBayNr.Item(i).Text) = 0) Then
              If i = 2 And (Not bBayCompeleted(1) And Len(ctlBayNr.Item(1).Text) > 0) Then
                Exit For
              End If
              mnuF10Accept.Enabled = True
              bDisableF10 = False
           End If
        Next i
  
        If Len(ctlBayNr.Item(0).Text) = 0 Then
           mnuF10Accept.Enabled = False
           bDisableF10 = True
        End If
       
        If Me.ctlBayNr.Item(0).Enabled Then
            If IsBlank(Me.ctlBayNr.Item(0).Text) Then
               mnuPrint.Enabled = False
            Else
               mnuPrint.Enabled = True
            End If
        End If
    Case mat_bay_validated
       mnuF2TrailerMessage.Enabled = bShowF2   ' If the Trailer/load has a msg
        mnuF3JobSchedules.Enabled = True
        mnuF4MultiBayDisplay.Enabled = False
        mnuF5QuickSort.Enabled = False
        mnuF7MultipleLoad.Enabled = bShowF7   ' If the current bay has multi-load
        mnuNPDLdsDisplay.Enabled = False
        mnuPDLdsDisplay.Enabled = False
        mnuSchByLds.Enabled = True
        mnuDolliesDisplay.Enabled = False   'True
        'kg disable - F10 button in the following cases below
        If bDuplicateLoad Then
           mnuF10Accept.Enabled = False
           bDisableF10 = True
        End If
   
        'kg - If any bay or dolly is not valid, disable F10 button
        For i = 0 To 2
           If (Not bBayCompeleted(i) And Len(ctlBayNr.Item(i).Text) > 0) Or (Not bDollyValidated(i) And Len(txtDolly.Item(i).Text) > 0) Then
              If (Not bDollyValidated(i) And Len(txtDolly.Item(i).Text) > 0) Or i = 0 Then
                mnuF10Accept.Enabled = False
                bDisableF10 = True
              End If
           ElseIf bBayCompeleted(0) And (Not bBayCompeleted(i) And Len(ctlBayNr.Item(i).Text) = 0) Then
              If i = 2 And (Not bBayCompeleted(1) And Len(ctlBayNr.Item(1).Text) > 0) Then
                Exit For
              End If
              mnuF10Accept.Enabled = True
              bDisableF10 = False
           End If
        Next i
   
        If Me.ctlBayNr.Item(0).Enabled Then
            If IsBlank(Me.ctlBayNr.Item(0).Text) Then
               mnuPrint.Enabled = False
               mnuF10Accept.Enabled = False
            Else
               mnuPrint.Enabled = True
            End If
        End If
    Case mat_dolly
      
        mnuF2TrailerMessage.Enabled = bShowF2  ' If the Trailer/load has a msg
        mnuF3JobSchedules.Enabled = True
        mnuF4MultiBayDisplay.Enabled = True
        mnuF5QuickSort.Enabled = True
        mnuF7MultipleLoad.Enabled = bShowF7 ' If the current bay has multi-load
        mnuNPDLdsDisplay.Enabled = True
        mnuPDLdsDisplay.Enabled = True
        mnuSchByLds.Enabled = True
        mnuDolliesDisplay.Enabled = True
'        mnuF10Accept.Enabled = True  'bAllowSave  'QS 2/1/05
        'kg disable - F10 button in the following cases below
        If bDuplicateLoad Then
           mnuF10Accept.Enabled = False
           bDisableF10 = True
        End If
   
        'kg - If any bay or dolly is not valid, disable F10 button
        For i = 0 To 2
           If (Not bBayCompeleted(i) And Len(ctlBayNr.Item(i).Text) > 0) Or (Not bDollyValidated(i) And Len(txtDolly.Item(i).Text) > 0) Then
              If (Not bDollyValidated(i) And Len(txtDolly.Item(i).Text) > 0) Then
                mnuF10Accept.Enabled = False
                bDisableF10 = True
              End If
           ElseIf bBayCompeleted(0) Then
              mnuF10Accept.Enabled = True
              bDisableF10 = False
           End If
        Next i
   
        If Len(ctlBayNr.Item(0).Text) = 0 And bPredispatchEmpty(0) = False And grdBayDetails(0).Columns(GRD1_COL_PD).Value = HFCS_INDICATOR_FALSE Then
           mnuF10Accept.Enabled = False
           bDisableF10 = True
        End If
        mnuPrint.Enabled = True
    End Select
   
    ' Set tool bar availability
    With ctlToolbarManager
        .buttonEnabled BT_F9_DOLLIES_ON_PROP, mnuDolliesDisplay.Enabled
        .buttonEnabled BT_F10_ACCEPT, mnuF10Accept.Enabled
        .buttonEnabled BT_F2_TLR_LD_MSG, mnuF2TrailerMessage.Enabled
        .buttonEnabled BT_F3_DRIVER_SCHED, mnuF3JobSchedules.Enabled
 
'<WR1024><ADID: dmx7plm> - Start
        .buttonEnabled BT_F4_MULTI_BAY, True
        .buttonEnabled BT_F5_QUICK_SEARCH, True
'<WR1024><ADID: dmx7plm> - End
        .buttonEnabled BT_F7_MULTI_LOAD, mnuF7MultipleLoad.Enabled
        .buttonEnabled BT_F8_PREDIS, mnuPDLdsDisplay.Enabled
       
        .buttonEnabled BT_S_F8_NON_PREDIS, mnuNPDLdsDisplay.Enabled
        .buttonEnabled BT_S_F3_SCHED_BY_LOAD, mnuSchByLds.Enabled
        .buttonEnabled BT_CTL_P_PRINT, mnuPrint.Enabled
    End With
   
'<WR1024><ADID: dmx7plm> - Start
        mnuF4MultiBayDisplay.Enabled = True
        mnuF5QuickSort.Enabled = True
'<WR1024><ADID: dmx7plm> - End
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetAvailMenuAndTool & Opt = " & Opt & " End")
    End If
Exit Sub
 
ErrorExit:
    glErrNum = 400
    gsError = "SetAvailMenuAndTool"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
    End If
End Sub
 
Private Sub DoDspSchedByDriver()
Dim vtemp() As Variant
Dim sStg As String
Dim dtSchedule As Date
If g_iDebug = 13 Then
   Call InfoLog("frmPDD DoDspSchedByDriver - Begin")
End If
 
'Calculate the schedule date
Select Case cboDow.Text
Case "MON"
     dtSchedule = DateAdd("d", -5, ctlWkendingDate.Value)
Case "TUE"
     dtSchedule = DateAdd("d", -4, ctlWkendingDate.Value)
Case "WED"
     dtSchedule = DateAdd("d", -3, ctlWkendingDate.Value)
Case "THU"
     dtSchedule = DateAdd("d", -2, ctlWkendingDate.Value)
Case "FRI"
     dtSchedule = DateAdd("d", -1, ctlWkendingDate.Value)
Case "SAT"
     dtSchedule = ctlWkendingDate.Value
Case "SUN"
      dtSchedule = DateAdd("d", -6, ctlWkendingDate.Value)
End Select
 sStg = txtJobNr.Text
vtemp = Array("PDD", _
                Trim$(sStg), _
                dtSchedule, _
                pInst.PDInfo.szPDJobDomSlic, _
                pInst.icdl)
oWSInfo.Notification_To_Start_Application _
         OGlobalFormNames.ScheduledInformation, _
         vtemp, _
         Me
If g_iDebug = 13 Then
   Call InfoLog("frmPDD DoDspSchedByDriver - End")
End If
End Sub
 
Private Sub DoDspSchedByLoadName()
' Shift-F3
Dim vtemp() As Variant
 
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoDspSchedByLoadName - Begin")
    End If
    If pInst.PDInfo.hPDHll.Count > 0 Then
       vtemp = Array("PDD", _
                   pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).szOrigin, _
                   pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).szOriginCn, _
                   pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).szOriginSrt, _
                   pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).szDestin, _
                   pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).szDestinCn, _
                   pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).szDestinSrt, _
                   0, vbNullString, vbNullString) ' pInst.PDInfo.hPDHll(pInst.PDInfo.hPDHll.Count).iSequenceNr, _
                   pInst.PDInfo.szPDEloc, _
                   pInst.PDInfo.szPDElocCn)
    Else
       vtemp = Array("PDD", "", "", "", "", "", "", 0, "", "")
    End If
 
    oWSInfo.Notification_To_Start_Application _
            OGlobalFormNames.SchedulesByLoadNameOutbound, _
            vtemp, _
            Me
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoDspSchedByLoadName - End")
    End If
Exit Sub
Error_Handler:
  glErrNum = 400
  gsError = "DoDspSchedByLoadName"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Sub
 
Private Sub DoMultiBayDisplay()
' F4
Dim lHwnd As Long 'WR01024, ADID dmx7plm
Dim vntTemp() As Variant
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoMultiBayDisplay - Begin")
    End If
    m_sChildName = OGlobalFormNames.MultiBaySearchF4
 
    gbProcessRetData = True
    'vntTemp = Array(vbNullString)
 
'<WR1024><ADID: dmx7plm><Date: 2nd March 2010> - Start
'<Need to enable F4 and F5 on load of screen 13>
    On Error Resume Next
    Me.Tag = Me.ActiveControl.Name
    If Me.Tag = "" Then Me.Tag = "txtJobNr"
'<WR1024><ADID: dmx7plm><Date: 2nd March 2010> - End
 
    If oWSNotify Is Nothing Then
        Set oWSNotify = New HFCSWSInformation.clsWSInfo
    End If
    Set oNotify = oWSNotify.NotificationClass
 
'<WR1024><ADID: dmx7plm><Date: 2nd March 2010> - Start
   
    bClickedF4 = True
    bClickedF5 = False
   
    lHwnd = GetWindowHandle(OGlobalFormNames.StartBaySearch)
    If (lHwnd = 0) Or (lHwnd <> 0 And Me.Tag <> "ctlBayNr") Then
        vntTemp = Array(vbNullString)
        oWSInfo.Notification_To_Start_Application _
                    OGlobalFormNames.MultiBaySearchF4, _
                    vntTemp, _
                    Me
        lHwnd = GetWindowHandle(OGlobalFormNames.StartBaySearch)
    End If
   
    If (lHwnd <> 0) Then
        If Me.Tag = "ctlBayNr" Then
            Call SendMessage(lHwnd, WM_MCL_CHANGEWRITE, 0, 0)
        Else
            Call SendMessage(lHwnd, WM_MCL_CHANGEREADONLY, 0, 0)
        End If
        Call ShowWindow(lHwnd, SW_SHOW)
    End If
    'Call SetForegroundWindow(lHwnd)
    Call SetActiveWindow(lHwnd)
'<WR1024><ADID: dmx7plm><Date: 2nd March 2010> - End
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoMultiBayDisplay - End")
    End If
    Exit Sub
Error_Handler:
  glErrNum = 400
  gsError = "DoMultiBayDisplay"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
End If
End Sub
 
Public Sub DoNotifyDriver()
Dim iRsp As Integer
Dim bRc As Boolean
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoNotifyDriver - Begin")
    End If
   
    '* * * * * * * * * * * * * * * * * * * * * * * * * *
    '   12/2000 - APP3DXH:  The users wanted a flag
    '   showing when the pre-dispatched load info was
    '   given to the driver.  This new db column is
    '   in the TFOPTPR table.
    '* * * * * * * * * * * * * * * * * * * * * * * * * *
    bRc = (chkDriverNotified.Value = vbUnchecked)
   
    If bRc Then
        iRsp = MsgBox(LoadResString(sTheDriverNotificationflagisnotset) & vbCr & _
                      LoadResString(sHaveYouNotifiedDriver), _
                  vbQuestion + vbYesNo + vbDefaultButton2, _
                  LoadResString(sNotifyDriver))
        If iRsp = vbYes Then
            pInst.PDInfo.szFdrDvrInfNtfIr = IR_TRUE
        Else
            pInst.PDInfo.szFdrDvrInfNtfIr = IR_FALSE
        End If
    Else
        pInst.PDInfo.szFdrDvrInfNtfIr = IR_TRUE
    End If
    chkDriverNotified.Value = IIf(pInst.PDInfo.szFdrDvrInfNtfIr = IR_TRUE, vbChecked, vbUnchecked)
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoNotifyDriver - End")
    End If
End Sub
 
Private Sub DoQuickSort()
On Error GoTo Error_Handler
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoQuickSort - Begin")
    End If
' F5
    Dim lHwnd As Long 'WR01024 ADID: dmx7plm
    Dim vntTemp() As Variant
    m_sChildName = OGlobalFormNames.QuickSearchF5
 
    gbProcessRetData = True
    vntTemp = Array(vbNullString)
 
'<WR1024><ADID: dmx7plm><Date: 2nd March 2010> - Start
'<Need to enable F4 and F5 on load of screen 13>
    On Error Resume Next
    Me.Tag = Me.ActiveControl.Name
    If Me.Tag = "" Then Me.Tag = "txtJobNr"
 
    If oWSNotify Is Nothing Then
        Set oWSNotify = New HFCSWSInformation.clsWSInfo
    End If
   
'<WR1024><ADID: dmx7plm><Date: 2nd March 2010> - Start
   
    bClickedF5 = True
    bClickedF4 = False
   
    lHwnd = GetWindowHandle(OGlobalFormNames.QuickSearchF5)
    If (lHwnd = 0) Or (lHwnd <> 0 And Me.Tag <> "ctlBayNr") Then
        vntTemp = Array(vbNullString)
        oWSInfo.Notification_To_Start_Application _
                    OGlobalFormNames.QuickSearchF5, _
                    vntTemp, _
                    Me
        lHwnd = GetWindowHandle(OGlobalFormNames.QuickSearchF5)
    End If
       
    If (lHwnd <> 0) Then
        If Me.Tag = "ctlBayNr" Then
            Call SendMessage(lHwnd, WM_MCL_CHANGEWRITE, 0, 0)
        Else
            Call SendMessage(lHwnd, WM_MCL_CHANGEREADONLY, 0, 0)
        End If
        Call ShowWindow(lHwnd, SW_SHOW)
    End If
    'Call SetForegroundWindow(lHwnd)
    Call SetActiveWindow(lHwnd)
'<WR1024><ADID: dmx7plm><Date: 2nd March 2010> - End
     If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoQuickSort - End")
    End If
    Exit Sub
Error_Handler:
  glErrNum = 400
  gsError = "DoQuickSort"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
  End If
End Sub
 
Private Sub DoDspMultiLd()
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoDspMultiLd - Begin")
    End If
    Dim vtemp() As Variant
   
    If oWSNotify Is Nothing Then
        Set oWSNotify = New HFCSWSInformation.clsWSInfo
    End If
  
    If pInst.PDInfo.hPDHll.Item(iBayIndex + 1).szMultLdIr = IR_TRUE Then
       vtemp = Array(1, pInst.PDInfo.hPDHll.Item(iBayIndex + 1).lCurFopTlrEntity)
       oWSInfo.Notification_To_Start_Application OGlobalFormNames.MultiLoads, vtemp, Me
    Else
       ctlWarningPopup.ShowWarning ctlBayNr(iBayIndex).hwnd, LoadResString(sThisisnotaMultipleLoad), 3000
    End If
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoDspMultiLd - End")
    End If
End Sub
 
Public Function GetPreDScreen(ValueType As Integer, sEloc As String, sDriver As String, sJobNr As String) As Boolean
' F8
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetPreDScreen - Begin")
    End If
    Dim vntTemp() As Variant
    m_sChildName = OGlobalFormNames.PreDispatchedLoads
   
    gbProcessRetData = True
   
    If oWSNotify Is Nothing Then
        Set oWSNotify = New HFCSWSInformation.clsWSInfo
    End If
   
    If (Len(grdBayDetails(0).Columns(4).Value) = 0) And _
       (g_bJobPredispatched = True Or g_bDriverPredispatched = True) Then
        If g_bJobPredispatched = True Then
            ReDim vntTemp(4)
            vntTemp(0) = Trim$(sJobNr)
            vntTemp(1) = Trim$(sDriver)
            If g_bELOCDiff = False Then
                vntTemp(3) = Trim$(sEloc)
            Else
                vntTemp(3) = vbNullString
            End If
            vntTemp(2) = NO_ELOC
        ElseIf g_bDriverPredispatched = True Then
            vntTemp = Array(ValueType, Trim$(sDriver), sEloc)
        End If
    Else
        vntTemp = Array(ValueType, sEloc, Trim$(sDriver))
    End If
   
'alters - db2 update upssys00.taplscr set  apl_fea_dsc_te = '13 - Arrival/Stand Alone Predispatch ' where apl_fea_nr = 13 and apl_nr = 3
    g_bJobPredispatched = False
    g_bDriverPredispatched = False
    g_bELOCDiff = False
   
    oWSInfo.Notification_To_Start_Application _
                        m_sChildName, _
                         vntTemp, _
                        Me
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetPreDScreen - End")
    End If
  Exit Function
Error_Handler:
  glErrNum = 400
  gsError = "GetPreDScreen"
  Call oProc.update_error_object(Me, gsError)
 
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
End If
End Function
 
Public Function GetNonPreDScreen(ValueData As String) As Boolean
' Shift-F8
    On Error GoTo Error_Handler
    Dim vntTemp() As Variant
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetNonPreDScreen - Begin")
    End If
    m_sChildName = OGlobalFormNames.NonPreDispatchedLoads
   
    gbProcessRetData = True
   
    
    If oWSNotify Is Nothing Then
        Set oWSNotify = New HFCSWSInformation.clsWSInfo
    End If
    vntTemp = Array(ValueData)
   
    oWSInfo.Notification_To_Start_Application _
                        OGlobalFormNames.NonPreDispatchedLoads, _
                        vntTemp, _
                        Me
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetNonPreDScreen - End")
    End If
    Exit Function
Error_Handler:
  glErrNum = 400
  gsError = "GetNonPreDScreen"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Function
 
Public Function GetDollyScreen(ValueData As String) As Boolean
' F9
    On Error GoTo Error_Handler
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetDollyScreen - Begin")
    End If
  
    Dim vntTemp() As Variant
 
    gbProcessRetData = True
    If oWSNotify Is Nothing Then
        Set oWSNotify = New HFCSWSInformation.clsWSInfo
    End If
    Set oNotify = oWSNotify.NotificationClass
    vntTemp = Array("F9", "R", "13")
    oWSInfo.Notification_To_Start_Application _
                OGlobalFormNames.EditDollies, _
                vntTemp, _
                Me
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD GetDollyScreen end")
    End If
   Exit Function
Error_Handler:
  glErrNum = 400
  gsError = "GetDollyScreen"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
End If
End Function
 
Private Sub DoSaveData()
On Error GoTo ErrorExit
Dim iRc As Integer
Dim i As Integer
Dim j As Integer
Dim tMsg As MSG
Dim oSendPredispatchtoIVIS As PredispatchObject.PreDispatch
Dim oWsInformation As HFCSWSInformation.clsWSInfo
 
'If bRoutingCanceled = True Then Exit Sub
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoSaveData - Begin")
    End If
bSaveData = True
Set oWsInformation = New HFCSWSInformation.clsWSInfo
    'clear keyboard buffer
    Do While PeekMessage(tMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
        'do nothing, just clear out buffer
    Loop
   
    'If there is a duplicate load, do not predispatch
    If bDuplicateLoad Then
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD DoSaveData duplicate load - End")
       End If
       Exit Sub
    End If
   
    'If any bay is not valid, do not predispatch
    For i = 0 To 2
       If Not bBayCompeleted(i) And Len(ctlBayNr.Item(i).Text) > 0 Then
          iPredispatch = i
          Exit For
       ElseIf Len(ctlBayNr.Item(i).Text) = 0 And Not bPredispatchEmpty(i) And Not bBayCompeleted(i) Then
          iPredispatch = i
          Exit For
       Else 'validate bay
      
       End If
    Next i
   
    If iPredispatch = 0 And i > 0 Then
       iPredispatch = i
    ElseIf iPredispatch = 0 Then
       If g_iDebug = 13 Then
           Call InfoLog("frmPDD DoSaveData iPredispatch = 0 - End bSaveData = False")
       End If
       bSaveData = False
       Exit Sub
    End If
   
    'If any dolly is not valid yet, do not predispatch
    For i = 0 To 2
       If Not bDollyValidated(i) And Len(txtDolly.Item(i).Text) > 0 Then
          'TSAT 3804 - try to validate dolly
          Call txtDolly_Validate(i, False)
          'if dolly is still not valid, exit.  Validate event will display error
          'saying dolly not valid
          If Not bDollyValidated(i) Then
              If g_iDebug = 13 Then
                  Call InfoLog("frmPDD DoSaveData bDollyValidated - End bSaveData = False")
              End If
              bSaveData = False
              Exit Sub
          End If
       End If
    Next i
     
    
    If Len(ctlBayNr.Item(0).Text) = 0 And Not bPredispatchEmpty(0) And Not bBayCompeleted(0) Then
       If g_iDebug = 13 Then
         Call InfoLog("frmPDD DoSaveData If Len(ctlBayNr.Item(0).Text) = 0 And Not bPredispatchEmpty(0) And Not bBayCompeleted(0) - End bSaveData = False")
       End If
       bSaveData = False
       Exit Sub
    End If
   
    'DoNotifyDriver
    pInst.PDInfo.szFdrDvrInfNtfIr = IR_TRUE
    ' This is what is doing the actual saving!
    iRc = UpdatePdDataNew()    'QS 2/9/2005
    
    If iRc <> DB_SUCCESS Then ' Save not successful
        'clear keyboard buffer
        Do While PeekMessage(tMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
            'do nothing, just clear out buffer
        Loop
       
        If iRc = DB_GRID_ERROR Then
            iPredispatch = 0
            bSaveData = False
            If g_iDebug = 13 Then
                  Call InfoLog("frmPDD DoSaveData bDollyValidated - End -  If iRc = DB_GRID_ERROR")
              End If
            Exit Sub
        ElseIf pInst.bArriveFlag Then
            ' Notify the calling program that a problem occured
            ' Close this program
           
            MsgBox LoadResString(sPreDispatchUnsuccessful)
           
            ' Return to the previous program
            bSaveData = False
            gbLoadPreDispatched = True
            bPreDispatched = True
            'For Arrival Tractor only
            'Sent back PDP info to Arrival Trackor Only screen.
            Unload Me
        Else
            MsgBox LoadResString(sPreDispatchUnsuccessful)
            bSaveData = False
           
            ' Close the program
            Unload Me
        End If
    Else ' Save went OK,PDP successful
        'clear keyboard buffer
        Do While PeekMessage(tMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
            'do nothing, just clear out buffer
        Loop
        If pInst.bArriveFlag Then
            MsgBox LoadResString(sPreDispatchsuccessful)
            Call WritePDPInfotoFeederShell
            bSaveData = False
            bPreDispatched = True
           
        Else
            'PreDispatch only
            Call WritePDPInfotoFeederShell
            MsgBox LoadResString(sPreDispatchsuccessful)
            bPreDispatched = True
            gbLoadPreDispatched = True
 
        End If
       
        If Not bFromGtF8 Then
            'send message to IVIS
            Set oSendPredispatchtoIVIS = New PredispatchObject.PreDispatch
   
            For i = 0 To 2
                If Not oOldColLoads(i, 0) Is Nothing Then
                    bPreDispatched = oSendPredispatchtoIVIS.SendNewPredispatch(oOldColLoads(i, 0), oOldColLoads(i, 1), oOldColLoads(i, 2), oOldColLoads(i, 0).PredispatchDate, oOldColLoads(i, 0).OutboundFeederScheduleWeekEndingDate, oOldColLoads(i, 0).JobNumber, oOldColLoads(i, 0).DriverName, oOldColLoads(i, 0).OutboundFeederScheduleDOW, _
                                                                               oOldColLoads(i, 0).TractorNumber, oOldColLoads(i, 0).PredispatchELOC, oOldColLoads(i, 0).PredispatchELOCCnyCd, Format$(oOldColLoads(i, 0).PredispatchDate, "mm/dd/yyyy"), Format$(oOldColLoads(i, 0).PredispatchTime, "hh:mm"), _
                                                                               oWsInformation.UserName, oWsInformation.ComputerName)
   
                    Sleep (500)
                End If
            Next i
            bPreDispatched = oSendPredispatchtoIVIS.SendNewPredispatch(oColLoads(0), oColLoads(1), oColLoads(2), Now, ctlWkendingDate.Value, txtJobNr.Text, txtDriverName.Text, CStr(cboDow.ListIndex), txtTractorNumber.Text, txtEloc.Text, sLocalCn, Format$(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy"), Format$(pInst.PDInfo.FeederScheduleEndTime, "hh:mm"), oWsInformation.UserName, oWsInformation.ComputerName)
            If Not oSendPredispatchtoIVIS Is Nothing Then
                oSendPredispatchtoIVIS.Dispose
            End If
   
            Set oSendPredispatchtoIVIS = Nothing
        End If
        bSaveData = False
        Set oWsInformation = Nothing
           
        Unload Me
 
    End If
 
    bSaveData = False
    Set oWsInformation = Nothing
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD DoSaveData - End")
    End If
Exit Sub
ErrorExit:
    glErrNum = 400
    gsError = "DoSaveData"
    Call oProc.update_error_object(Me, gsError)
    Screen.MousePointer = vbNormal
    bSaveData = False
     If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
      '  Set oSendPredispatchtoIVIS = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
End Sub
 
 
Private Sub ShowLdTlrMsg()
On Error GoTo ErrorHandler
'
Dim sTlrNo As String
Dim vntTemp() As Variant
Dim bLoadreadonly As Boolean
Dim bTrilerreadonly As Boolean
 
    If g_iDebug = 13 Then
      Call InfoLog("frmPDD ShowLdTlrMsg - Begin")
    End If
 
    Screen.MousePointer = vbHourglass
       
    ' check trailer message
   
       
    'Call new TrailLoadScreen which is readonly  QS 12/08/04
    If pInst.PDInfo.hPDHll(ActiveControl.Index + 1).lLdMsgId > 0 Or _
       Len(pInst.PDInfo.hPDHll(ActiveControl.Index + 1).szBayNr) > 0 Then
        m_bEditTrailerMessage = True
        m_bEditLoadMessage = True
    End If
    If pInst.PDInfo.hPDHll(ActiveControl.Index + 1).lTlrMsgId > 0 Or _
       Len(pInst.PDInfo.hPDHll(ActiveControl.Index + 1).szBayNr) > 0 Then
        m_bEditTrailerMessage = True
        m_bEditLoadMessage = True
    End If
    vntTemp = Array(PopulateLoad(pInst.PDInfo.hPDHll(ActiveControl.Index + 1)), m_bEditLoadMessage, m_bEditTrailerMessage)
   
    oWSInfo.Notification_To_Start_Application OGlobalFormNames.TrailerLoadMessages, vntTemp, Me
 
   
    Screen.MousePointer = vbDefault
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ShowLdTlrMsg - End")
    End If
    Exit Sub
ErrorHandler:
glErrNum = 400
gsError = "ShowLdTlrMsg"
Call oProc.update_error_object(Me, gsError)
If oErrorObject.error_routine(oEventlog.FeederShell, _
                               IIf(Err.Number <> 0, Err.Number, glErrNum), _
                               Err.Description & " Module:" & gsError, _
                               oProc, _
                               ERROR_MSG, _
                               FEEDER_DISPATCH_DRIVER) Then
    Set oEventlog = Nothing
    MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
   ' Unload Me
End If
End Sub
 
Private Sub Print_to_Printer()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD Print_to_Printer - Begin")
    End If
    Print_records ToPrinter
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD Print_to_Printer - End")
    End If
End Sub
 
 
Private Function Print_records(Destination As PrintDestination)
On Error GoTo Error_Handler
Dim oPrint As HFCSPrint.clsComm
Dim rs As ADODB.Recordset
Dim sPrintHead As String
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD Print_records - Begin")
    End If
 
    Set oPrint = New HFCSPrint.clsComm
    Set rs = New ADODB.Recordset
    sPrintHead = LoadResString(sPredispatchDriver)
    CreatePrintRecordSet rs
    oPrint.Print_Recordset sPrintHead, rs, Destination, Landscape
    rs.Close
    Set oPrint = Nothing
    Set rs = Nothing
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD Print_records - End")
    End If
    Exit Function
   
Error_Handler:
glErrNum = 400
gsError = "Print_records"
Call oProc.update_error_object(Me, gsError)
 
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                                 IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                 Err.Description & " Module:" & gsError, _
                                 oProc, _
                                 ERROR_MSG, _
                                 FEEDER_DISPATCH_DRIVER) Then
      Set oEventlog = Nothing
      MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '  Unload Me
End If
End Function
 
Private Sub CreatePrintRecordSet(ByRef rs As ADODB.Recordset)
On Error GoTo Error_Handler
Dim i         As Integer
Dim sHeads()  As String
Dim iWidths() As Integer
Dim iTotHeads As Integer: iTotHeads = 23
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CreatePrintRecordSet - Begin")
    End If
   
    If rs.State <> adStateClosed Then
        rs.Close
    End If
   
    ReDim sHeads(iTotHeads)
    ReDim iWidths(iTotHeads)
   
    sHeads(1) = LoadResString(233)
    sHeads(2) = LoadResString(567)
    sHeads(3) = LoadResString(234)
    sHeads(4) = LoadResString(466)
    sHeads(5) = LoadResString(231)
    sHeads(6) = LoadResString(sRES_BAY_NUM)
    sHeads(7) = LoadResString(sRES_TRAILER) & " " & LoadResString(sRES_TRAILER_NUM)
    sHeads(8) = LoadResString(210)
    sHeads(9) = LoadResString(sRES_ORIG)
    sHeads(10) = LoadResString(sRES_ORIG_SORT) & " "
    sHeads(11) = LoadResString(sRES_DEST)
    sHeads(12) = LoadResString(sRES_DEST_SRT)
    sHeads(13) = LoadResString(sRES_LD_SEQ)
    sHeads(14) = LoadResString(sLC)
    sHeads(15) = LoadResString(sRES_RT)
    sHeads(16) = LoadResString(sRES_PCS)
    sHeads(17) = LoadResString(sRES_PCT)
    sHeads(18) = LoadResString(sRES_TAG)
    sHeads(19) = LoadResString(1095)
    sHeads(20) = LoadResString(232)
    sHeads(21) = LoadResString(sRES_DEP_TIME)
    sHeads(22) = LoadResString(sRES_HAZMAT)
    sHeads(23) = LoadResString(sRES_REMARKS)
   
    iWidths(1) = 6
    iWidths(2) = 15
    iWidths(3) = 9
    iWidths(4) = 5
    iWidths(5) = 2
    iWidths(6) = 4
    iWidths(7) = 12
    iWidths(8) = 3
    iWidths(9) = 5
    iWidths(10) = 1
    iWidths(11) = 5
    iWidths(12) = 1
    iWidths(13) = 2
    iWidths(14) = 1
    iWidths(15) = 2
    iWidths(16) = 7
    iWidths(17) = 3
    iWidths(18) = 3
    iWidths(19) = 4
    iWidths(20) = 5
    iWidths(21) = 9
    iWidths(22) = 3
    iWidths(23) = 32
 
    With rs.Fields
        For i = 1 To iTotHeads
            .Append sHeads(i), adVarChar, iWidths(i), adFldIsNullable
        Next i
    End With
    rs.Open
   
    ' Set base data only for the first record
    rs.AddNew
    With pInst.PDInfo
        rs.Fields(0).Value = .szPDJobNr   '  Job #
        rs.Fields(1).Value = Left$(.szPDDvrNa, iWidths(2))   '  Driver name
        rs.Fields(2).Value = .szPDTrcNr     ' Veh #
        rs.Fields(3).Value = .szPDEloc     ' ELoc
        rs.Fields(4).Value = IIf(.szFdrDvrInfNtfIr = IR_TRUE, "Yes", "No")     ' Dvr Notf
    End With
   
    ' Set the load/trailer/bay details
    For i = 1 To pInst.PDInfo.hPDHll.Count
        If i > 1 Then
            rs.AddNew
            rs.Fields(0).Value = ""
            rs.Fields(1).Value = ""
            rs.Fields(2).Value = ""
            rs.Fields(3).Value = ""
            rs.Fields(4).Value = ""
        End If
        With pInst.PDInfo.hPDHll(i)
            rs.Fields(5) = .szBayNr  ' "Bay #"
            rs.Fields(6) = .szTlrNr  ' "Trailer #"
            rs.Fields(7) = .szEqpTyp ' "Typ"
            rs.Fields(8) = .szOrigin   ' "Orig"
            rs.Fields(9) = .szOriginSrt  ' "OS"
            rs.Fields(10) = .szDestin ' "Dest"
            rs.Fields(11) = .szDestinSrt ' "DS"
            rs.Fields(12) = .iSequenceNr ' "Sq"
            rs.Fields(13) = .szLdCd ' "LC"
            rs.Fields(14) = .szRtId  ' "Rt"
            rs.Fields(15) = .lPieces ' "Pcs"
            rs.Fields(16) = .iPercent ' "%"
            rs.Fields(17) = .szTgCd ' Tag"
            rs.Fields(18) = .szDspCrtDt  ' "CrDt"
            rs.Fields(19) = .szDspDueDt  ' "DueDt"
            rs.Fields(20) = .szDspScdDptTm ' "SchDepTim"
            rs.Fields(21) = .szHzMt  ' "Haz"
            rs.Fields(22) = .szRemarks ' "Remarks"
            rs.Update
        End With
    Next i
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CreatePrintRecordSet - Begin")
    End If
    Exit Sub
Error_Handler:
glErrNum = 400
gsError = "CreatePrintRecordSet"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                        IIf(Err.Number <> 0, Err.Number, glErrNum), _
                        Err.Description & " Module:" & gsError, _
                        oProc, _
                        ERROR_MSG, _
                        FEEDER_DISPATCH_DRIVER) Then
Set oEventlog = Nothing
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'Unload Me
End If
End Sub
 
Public Property Get ElocValid() As Boolean
    If g_iDebug = 13 Then
       Call InfoLog("Get ElocValid - Begin")
    End If
    ElocValid = mbElocValid
    If g_iDebug = 13 Then
       Call InfoLog("Get ElocValid - End")
    End If
End Property
 
Public Property Let ElocValid(ByVal vNewValue As Boolean)
    If g_iDebug = 13 Then
       Call InfoLog("let ElocValid - Begin")
    End If
   
    mbElocValid = vNewValue
   
    If g_iDebug = 13 Then
       Call InfoLog("let ElocValid - End")
    End If
   
End Property
 
Public Property Get ClearData() As Boolean
    If g_iDebug = 13 Then
       Call InfoLog("Get ClearData - Begin")
    End If
   
    ClearData = mbClearData
   
    If g_iDebug = 13 Then
       Call InfoLog("Get ClearData - End")
    End If
End Property
 
Public Property Let ClearData(ByVal vNewValue As Boolean)
    If g_iDebug = 13 Then
       Call InfoLog("let ClearData - Begin")
    End If
    mbClearData = vNewValue
    If g_iDebug = 13 Then
       Call InfoLog("let ClearData - End")
    End If
End Property
 
Private Sub WritePDPInfotoFeederShell() 'QS. 11/24/04
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD WritePDPInfotoFeederShell - Begin")
    End If
    On Error GoTo Error_Handler
       
        Dim oWsInformation As HFCSWSInformation.clsWSInfo
        Dim sMessage1 As String
        Dim sMessage2 As String
        Dim sMessage3 As String
        Dim sBayNum As String
        Dim sTrailerNum As String
        Dim i As Integer
        Set oWsInformation = New HFCSWSInformation.clsWSInfo
 
        Dim sMessage As String
        Dim oScratchPad As HFCSScratchPadClient.clsScratchPad
        Dim sLoadName As String
       
        Set oScratchPad = New HFCSScratchPadClient.clsScratchPad
        For i = 0 To 2
            sRouteCodeMessage(i) = ""
            sRouteCodeChanged(i) = ""
        Next
        If iPredispatch = 1 Then
           sTrailerNum = " @@" & pInst.PDInfo.hPDHll.Item(1).szTlrNr  ' WR00017
           sBayNum = " @@" & pInst.PDInfo.hPDHll.Item(1).szBayNr
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(1).szTlrNr)) > 0 Then
               sMessage = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(1).szTlrNr & " " & pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
               ' Format Route Code Message
               If pInst.PDInfo.hPDHll.Item(1).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
                  sRouteCodeMessage(0) = LoadResString(1114)
                   
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%2", "Number")
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%3", pInst.PDInfo.hPDHll.Item(1).szTlrNr)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%4", sLoadName)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%5", pInst.PDInfo.szPDEloc)
               End If
              
               If pInst.PDInfo.hPDHll.Item(1).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(1).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(0) = LoadResString(1115)
              
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%2", "number")
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%3", pInst.PDInfo.hPDHll.Item(1).szTlrNr)
               End If
           Else
               sMessage = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(1).szEqpTyp & " " & pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
               ' Format Route Code Message
               If pInst.PDInfo.hPDHll.Item(1).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
                  sRouteCodeMessage(0) = LoadResString(1114)
                   
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%2", "Type")
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%3", pInst.PDInfo.hPDHll.Item(1).szEqpTyp)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%4", sLoadName)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%5", pInst.PDInfo.szPDEloc)
               End If
              
               If pInst.PDInfo.hPDHll.Item(1).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(1).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(0) = LoadResString(1115)
 
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%2", "Type")
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%3", pInst.PDInfo.hPDHll.Item(1).szEqpTyp)
               End If
           End If
        ElseIf iPredispatch = 2 Then
           sTrailerNum = " @@" & pInst.PDInfo.hPDHll.Item(1).szTlrNr & " @@" & pInst.PDInfo.hPDHll.Item(2).szTlrNr
           sBayNum = " @@" & pInst.PDInfo.hPDHll.Item(1).szBayNr & " @@" & pInst.PDInfo.hPDHll.Item(2).szBayNr
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(1).szTlrNr)) > 0 Then
               sMessage1 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(1).szTlrNr & " " & pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
               ' Format Route Code Message
               If pInst.PDInfo.hPDHll.Item(1).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
                  sRouteCodeMessage(0) = LoadResString(1114)
                   
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%2", "Number")
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%3", pInst.PDInfo.hPDHll.Item(1).szTlrNr)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%4", sLoadName)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%5", pInst.PDInfo.szPDEloc)
               End If
          
               If pInst.PDInfo.hPDHll.Item(1).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(1).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(0) = LoadResString(1115)
                 
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%2", "Number")
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%3", pInst.PDInfo.hPDHll.Item(1).szTlrNr)
               End If
           Else
               sMessage1 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(1).szEqpTyp & " " & pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
               ' Format Route Code Message
               If pInst.PDInfo.hPDHll.Item(1).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
                  sRouteCodeMessage(0) = LoadResString(1114)
                   
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%2", "Type")
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%3", pInst.PDInfo.hPDHll.Item(1).szEqpTyp)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%4", sLoadName)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%5", pInst.PDInfo.szPDEloc)
               End If
              
               If pInst.PDInfo.hPDHll.Item(1).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(1).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(0) = LoadResString(1115)
 
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%2", "Type")
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%3", pInst.PDInfo.hPDHll.Item(1).szEqpTyp)
               End If
           End If
          
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(2).szTlrNr)) > 0 Then
               sMessage2 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(2).szTlrNr & " " & pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
               If pInst.PDInfo.hPDHll.Item(2).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
                  sRouteCodeMessage(1) = LoadResString(1114)
                   
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%2", "Number")
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%3", pInst.PDInfo.hPDHll.Item(2).szTlrNr)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%4", sLoadName)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%5", pInst.PDInfo.szPDEloc)
               End If
              
               If pInst.PDInfo.hPDHll.Item(2).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(2).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(1) = LoadResString(1115)
                 
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%2", "Number")
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%3", pInst.PDInfo.hPDHll.Item(2).szTlrNr)
               End If
           Else
               sMessage2 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(2).szEqpTyp & " " & pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
               If pInst.PDInfo.hPDHll.Item(2).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
                  sRouteCodeMessage(1) = LoadResString(1114)
                   
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%2", "Type")
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%3", pInst.PDInfo.hPDHll.Item(2).szEqpTyp)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%4", sLoadName)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%5", pInst.PDInfo.szPDEloc)
               End If
               If pInst.PDInfo.hPDHll.Item(2).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(2).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(1) = LoadResString(1115)
                  
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%2", "Type")
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%3", pInst.PDInfo.hPDHll.Item(2).szEqpTyp)
               End If
           End If
           
           sMessage = sMessage1 & " @@" & sMessage2
        Else
           sTrailerNum = " @@" & pInst.PDInfo.hPDHll.Item(1).szTlrNr & " @@" & pInst.PDInfo.hPDHll.Item(2).szTlrNr & " @@" & pInst.PDInfo.hPDHll.Item(3).szTlrNr
           sBayNum = " @@" & pInst.PDInfo.hPDHll.Item(1).szBayNr & " @@" & pInst.PDInfo.hPDHll.Item(2).szBayNr & " @@" & pInst.PDInfo.hPDHll.Item(3).szBayNr
 
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(1).szTlrNr)) > 0 Then
               sMessage1 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(1).szTlrNr & " " & pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
               ' Format Route Code Message
               If pInst.PDInfo.hPDHll.Item(1).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
                  sRouteCodeMessage(0) = LoadResString(1114)
                   
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%2", "Number")
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%3", pInst.PDInfo.hPDHll.Item(1).szTlrNr)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%4", sLoadName)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%5", pInst.PDInfo.szPDEloc)
               End If
              If pInst.PDInfo.hPDHll.Item(1).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(1).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(0) = LoadResString(1115)
                 
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%2", "Number")
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%3", pInst.PDInfo.hPDHll.Item(1).szTlrNr)
               End If
           Else
               sMessage1 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(1).szEqpTyp & " " & pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
               ' Format Route Code Message
               If pInst.PDInfo.hPDHll.Item(1).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(1).szOrigin & " " & pInst.PDInfo.hPDHll.Item(1).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(1).szDestin & " " & pInst.PDInfo.hPDHll.Item(1).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(1).iSequenceNr
                  sRouteCodeMessage(0) = LoadResString(1114)
                   
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%2", "Type")
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%3", pInst.PDInfo.hPDHll.Item(1).szEqpTyp)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%4", sLoadName)
                  sRouteCodeMessage(0) = Replace$(sRouteCodeMessage(0), "%5", pInst.PDInfo.szPDEloc)
               End If
               If pInst.PDInfo.hPDHll.Item(1).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(1).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(0) = LoadResString(1115)
                   
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%1", pInst.PDInfo.hPDHll.Item(1).szRtId)
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%2", "Type")
                  sRouteCodeChanged(0) = Replace$(sRouteCodeChanged(0), "%3", pInst.PDInfo.hPDHll.Item(1).szEqpTyp)
              End If
           End If
          
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(2).szTlrNr)) > 0 Then
               sMessage2 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(2).szTlrNr & " " & pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
               If pInst.PDInfo.hPDHll.Item(2).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
                  sRouteCodeMessage(1) = LoadResString(1114)
                   
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%2", "Number")
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%3", pInst.PDInfo.hPDHll.Item(2).szTlrNr)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%4", sLoadName)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%5", pInst.PDInfo.szPDEloc)
               End If
          
               If pInst.PDInfo.hPDHll.Item(2).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(2).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(1) = LoadResString(1115)
                 
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%2", "Number")
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%3", pInst.PDInfo.hPDHll.Item(2).szTlrNr)
               End If
           Else
               sMessage2 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(2).szEqpTyp & " " & pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
               If pInst.PDInfo.hPDHll.Item(2).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(2).szOrigin & " " & pInst.PDInfo.hPDHll.Item(2).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(2).szDestin & " " & pInst.PDInfo.hPDHll.Item(2).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(2).iSequenceNr
                  sRouteCodeMessage(1) = LoadResString(1114)
                   
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%2", "Type")
                 sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%3", pInst.PDInfo.hPDHll.Item(2).szEqpTyp)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%4", sLoadName)
                  sRouteCodeMessage(1) = Replace$(sRouteCodeMessage(1), "%5", pInst.PDInfo.szPDEloc)
               End If
               If pInst.PDInfo.hPDHll.Item(2).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(2).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(1) = LoadResString(1115)
                  
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%1", pInst.PDInfo.hPDHll.Item(2).szRtId)
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%2", "Type")
                  sRouteCodeChanged(1) = Replace$(sRouteCodeChanged(1), "%3", pInst.PDInfo.hPDHll.Item(2).szEqpTyp)
              End If
           End If
          
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(3).szTlrNr)) > 0 Then
               sMessage3 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(3).szTlrNr & "," & pInst.PDInfo.hPDHll.Item(3).szOrigin & " " & pInst.PDInfo.hPDHll.Item(3).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(3).szDestin & " " & pInst.PDInfo.hPDHll.Item(3).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(3).iSequenceNr
               If pInst.PDInfo.hPDHll.Item(3).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(3).szOrigin & " " & pInst.PDInfo.hPDHll.Item(3).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(3).szDestin & " " & pInst.PDInfo.hPDHll.Item(3).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(3).iSequenceNr
                  sRouteCodeMessage(2) = LoadResString(1114)
                   
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%1", pInst.PDInfo.hPDHll.Item(3).szRtId)
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%2", "Number")
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%3", pInst.PDInfo.hPDHll.Item(3).szTlrNr)
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%4", sLoadName)
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%5", pInst.PDInfo.szPDEloc)
               End If
          
               If pInst.PDInfo.hPDHll.Item(3).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(3).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(2) = LoadResString(1115)
                 
                  sRouteCodeChanged(2) = Replace$(sRouteCodeChanged(2), "%1", pInst.PDInfo.hPDHll.Item(3).szRtId)
                  sRouteCodeChanged(2) = Replace$(sRouteCodeChanged(2), "%2", "Number")
                  sRouteCodeChanged(2) = Replace$(sRouteCodeChanged(2), "%3", pInst.PDInfo.hPDHll.Item(3).szTlrNr)
               End If
           Else
               sMessage3 = LoadResString(1102) & pInst.PDInfo.hPDHll.Item(3).szEqpTyp & " " & pInst.PDInfo.hPDHll.Item(3).szOrigin & " " & pInst.PDInfo.hPDHll.Item(3).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(3).szDestin & " " & pInst.PDInfo.hPDHll.Item(3).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(3).iSequenceNr
               If pInst.PDInfo.hPDHll.Item(3).InvalidLoadRouting = True Then
                  sLoadName = pInst.PDInfo.hPDHll.Item(3).szOrigin & " " & pInst.PDInfo.hPDHll.Item(3).szOriginSrt & " " & pInst.PDInfo.hPDHll.Item(3).szDestin & " " & pInst.PDInfo.hPDHll.Item(3).szDestinSrt & " " & pInst.PDInfo.hPDHll.Item(3).iSequenceNr
                  sRouteCodeMessage(2) = LoadResString(1114)
                   
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%1", pInst.PDInfo.hPDHll.Item(3).szRtId)
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%2", "Number")
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%3", pInst.PDInfo.hPDHll.Item(3).szEqpTyp)
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%4", sLoadName)
                  sRouteCodeMessage(2) = Replace$(sRouteCodeMessage(2), "%5", pInst.PDInfo.szPDEloc)
               End If
               If pInst.PDInfo.hPDHll.Item(3).SendPredispatchPreforecast = True Or _
                  pInst.PDInfo.hPDHll.Item(3).EmptyRouteCodeChanged = True Then
                  sRouteCodeChanged(2) = LoadResString(1115)
                   
                  sRouteCodeChanged(2) = Replace$(sRouteCodeChanged(2), "%1", pInst.PDInfo.hPDHll.Item(3).szRtId)
                  sRouteCodeChanged(2) = Replace$(sRouteCodeChanged(2), "%2", "Type")
                  sRouteCodeChanged(2) = Replace$(sRouteCodeChanged(2), "%3", pInst.PDInfo.hPDHll.Item(3).szEqpTyp)
              End If
           End If
          
           sMessage = sMessage1 & " @@" & sMessage2 & " @@" & sMessage3
        End If
        PDPInfo2 = sMessage
        If iPredispatch = 1 Then
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(1).szTlrNr)) > 0 Then
               sMessage = LoadResString(1103) & pInst.PDInfo.szPDDvrNa & "," & pInst.PDInfo.szPDJobNr & LoadResString(1104) & pInst.PDInfo.hPDHll.Item(1).szTlrNr & " ELOC: " & pInst.PDInfo.szPDEloc
           Else
               sMessage = LoadResString(1103) & pInst.PDInfo.szPDDvrNa & "," & pInst.PDInfo.szPDJobNr & LoadResString(1104) & pInst.PDInfo.hPDHll.Item(1).szEqpTyp & " ELOC: " & pInst.PDInfo.szPDEloc
           End If
        ElseIf iPredispatch = 2 Then
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(1).szTlrNr)) > 0 Then
               sMessage = LoadResString(1103) & pInst.PDInfo.szPDDvrNa & "," & pInst.PDInfo.szPDJobNr & LoadResString(1104) & pInst.PDInfo.hPDHll.Item(1).szTlrNr & ", "
           Else
               sMessage = LoadResString(1103) & pInst.PDInfo.szPDDvrNa & "," & pInst.PDInfo.szPDJobNr & LoadResString(1104) & pInst.PDInfo.hPDHll.Item(1).szEqpTyp & ", "
           End If
          
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(2).szTlrNr)) > 0 Then
               sMessage = sMessage & pInst.PDInfo.hPDHll.Item(2).szTlrNr & " ELOC: " & pInst.PDInfo.szPDEloc
           Else
               sMessage = sMessage & pInst.PDInfo.hPDHll.Item(2).szEqpTyp & " ELOC: " & pInst.PDInfo.szPDEloc
           End If
        Else
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(1).szTlrNr)) > 0 Then
               sMessage = LoadResString(1103) & pInst.PDInfo.szPDDvrNa & "," & pInst.PDInfo.szPDJobNr & LoadResString(1104) & pInst.PDInfo.hPDHll.Item(1).szTlrNr & ", "
           Else
               sMessage = LoadResString(1103) & pInst.PDInfo.szPDDvrNa & "," & pInst.PDInfo.szPDJobNr & LoadResString(1104) & pInst.PDInfo.hPDHll.Item(1).szEqpTyp & ", "
           End If
       
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(2).szTlrNr)) > 0 Then
               sMessage = sMessage & pInst.PDInfo.hPDHll.Item(2).szTlrNr & ", "
           Else
               sMessage = sMessage & pInst.PDInfo.hPDHll.Item(2).szEqpTyp & ", "
           End If
          
           If Len(Trim$(pInst.PDInfo.hPDHll.Item(3).szTlrNr)) > 0 Then
               sMessage = sMessage & pInst.PDInfo.hPDHll.Item(3).szTlrNr & " ELOC: " & pInst.PDInfo.szPDEloc
           Else
               sMessage = sMessage & pInst.PDInfo.hPDHll.Item(3).szEqpTyp & " ELOC: " & pInst.PDInfo.szPDEloc
           End If
        End If
        PDPInfo1 = sMessage
        sMessage = PDPInfo1 & " @@ " & PDPInfo2
        If Not bFromGtF8 Then  'When data from GT F8, Arrival screen will write this info to scratch pad.
           oScratchPad.AddScratchPadRecord 9999, "", False, ScratchPadMsgTypes.IPLD, NonInitMsg, "", FeederScratchPad, 1025, sMessage, , oWsInformation.HFCSUserName, , sBayNum, sTrailerNum
           For i = 0 To iPredispatch
               If Len(Trim$(sRouteCodeMessage(i))) > 0 Then
                   oScratchPad.AddScratchPadRecord 9999, "", False, ScratchPadMsgTypes.IPLD, NonInitMsg, "", FeederScratchPad, 1025, sRouteCodeMessage(i), , oWsInformation.HFCSUserName, , pInst.PDInfo.hPDHll.Item(i + 1).szBayNr, pInst.PDInfo.hPDHll.Item(i + 1).szTlrNr
               End If
              
               If Len(Trim$(sRouteCodeChanged(i))) > 0 Then
                   oScratchPad.AddScratchPadRecord 9999, "", False, ScratchPadMsgTypes.IPLD, NonInitMsg, "", FeederScratchPad, 1025, sRouteCodeChanged(i), , oWsInformation.HFCSUserName, , pInst.PDInfo.hPDHll.Item(i + 1).szBayNr, pInst.PDInfo.hPDHll.Item(i + 1).szTlrNr
               End If
           Next
        End If
        Set oScratchPad = Nothing
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD WritePDPInfotoFeederShell - End")
    End If
    Exit Sub
Error_Handler:
glErrNum = 400
gsError = "WritePDPInfotoFeederShell"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                        IIf(Err.Number <> 0, Err.Number, glErrNum), _
                        Err.Description & " Module:" & gsError, _
                        oProc, _
                        ERROR_MSG, _
                        FEEDER_DISPATCH_DRIVER) Then
Set oEventlog = Nothing
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
Resume Next
End If
End Sub
 
Public Function TrlLoadMsgScrFinished(curLoad As GLOBALDEFS.HFCSLOAD, _
                                      LDStatus As MSGSTATUS, _
                                      sLdMsg As String, _
                                      TRLStatus As MSGSTATUS, _
                                      sTrlMsg As String) As Boolean
    On Error GoTo ErrorHandler
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD TrlLoadMsgScrFinished - Begin")
    End If
    Set oLoad = New HFCSLoadObject.HFCSLOAD
    oLoad.SpecifyConnection oClsDb.DBClass
   
    If oLoad.GetLoadByEntityKey(pInst.PDInfo.hPDHll.Item(iBayIndex + 1).lCurFopTlrEntity, False) = GET_BY_ENTITY_KEY_FAILURE Then
       MsgBox "Not OK"
    End If
   
 
    oLoad.LoadMessage = sLdMsg
    oLoad.TrailerMessage = sTrlMsg
    sTrailerMsg = sTrlMsg
    sLoadMsg = sLdMsg
   
    oLoad.Update
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD TrlLoadMsgScrFinished - End")
    End If
   
    Exit Function
   
ErrorHandler:
glErrNum = 400
gsError = "TrlLoadMsgScrFinished"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                        IIf(Err.Number <> 0, Err.Number, glErrNum), _
                        Err.Description & " Module:" & gsError, _
                        oProc, _
                        ERROR_MSG, _
                        FEEDER_DISPATCH_DRIVER) Then
Set oEventlog = Nothing
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'Unload Me
End If
End Function
'===================================================================
' Sub:          setDowAndWkEndingDate()
' Description:  set days of the week description list to cboDOW list
'               and set default dow and wk ending date.
'===================================================================
Private Sub setDowAndWkEndingDate()
    Dim dtCurrent As Date
    Dim rsDowList As ADODB.Recordset
    Dim i As Integer
    Dim oSortCal As HFCSSortsObject.clsSortCalendar
    Dim sSortEnd As String
    Dim vSortDate As Variant
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD setDowAndWkEndingDate - Begin")
    End If
    Set oSortCal = New HFCSSortsObject.clsSortCalendar
    Set rsDowList = oClsDb.getDowListsRecordset
    If rsDowList Is Nothing Then
         setDOWList
    ElseIf rsDowList.EOF Then
        setDOWList
        rsDowList.Close
        Set rsDowList = Nothing
    Else
        rsDowList.MoveFirst
        For i = 0 To 6
            g_sWeekCode(i) = Trim$(rsDowList.Fields("DOW_CD_DSC_TE").Value)
            rsDowList.MoveNext
            cboDow.AddItem g_sWeekCode(i)
        Next i
        rsDowList.Close
        Set rsDowList = Nothing
    End If
   
    'get the current operational day
    Call oSortCal.SpecifyConnection(oClsDb.DBClass)
    vSortDate = oSortCal.SortInfo
    ' TT02142 - GRENC - 07.083.0026 issue with predispatch (13) screen pulling up wrong date.
'    dtCurrent = vSortDate(1)
'    'if time is not between times defined in the sort calendar,
'    'use operational date based off of reset time.
'    If Not (DateDiff("s", vSortDate(4), Now) >= 0 _
'          And DateDiff("s", Now, vSortDate(5)) >= 0) Then
'        dtCurrent = vSortDate(6)
'    End If
    dtCurrent = CStr(vSortDate(enumSortInfoResults.SI_OPERATIONAL_DATE))
    ' End - TT02142
   
    g_sDow = Weekday(dtCurrent)
    cboDow.Text = g_sWeekCode(g_sDow - 1)
    ' find this week's week ending date
    ctlWkendingDate.Value = Format(DateAdd("d", 7 - Weekday(dtCurrent), dtCurrent), "Short Date") 'Date), "Short Date")
 
    Set oSortCal = Nothing
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD setDowAndWkEndingDate - End")
    End If
End Sub
'=======================================================================
' Function:         setDOWList()
' Description:      set days of the week description list to g_sWeekCode.
'                   The function is called when TVALDOW table is not
'                   available.
'=======================================================================
Private Sub setDOWList()
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD setDOWList - Begin")
        End If
        g_sWeekCode(0) = LoadResString(705)  '"SUN"
        g_sWeekCode(1) = LoadResString(706)  '"MON"
        g_sWeekCode(2) = LoadResString(707)  '"TUE"
        g_sWeekCode(3) = LoadResString(708)  '"WED"
        g_sWeekCode(4) = LoadResString(709)  '"THU"
        g_sWeekCode(5) = LoadResString(710)  '"FRI"
        g_sWeekCode(6) = LoadResString(711)  '"SAT"
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD setDOWList - End")
        End If
End Sub
Public Function SelectedDolly(sSelectedDolly As String)
On Error GoTo Error_Handler
 
 
Dim sDollyInfo As Variant
    If g_iDebug = 13 Then
       Call InfoLog("SelectedDolly - Begin")
    End If
   
   If ctlWarningPopup.WarningVisible = True Then
      ctlWarningPopup.ClearWarning
   End If
  
   sDollyInfo = Split(sSelectedDolly, "|")   'QS 1/7/05
   m_sDllyEntityKey(iDollyIndex) = sDollyInfo(0)
   pInst.PDInfo.hPDHll.Item(iDollyIndex + 1).lCurFopDolEntity = m_sDllyEntityKey(iDollyIndex)
   txtDolly.Item(iDollyIndex).Text = sDollyInfo(2)
    If g_iDebug = 13 Then
       Call InfoLog("SelectedDolly - End")
    End If
  Exit Function
Error_Handler:
glErrNum = 400
gsError = "SelectedDolly"
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
End If
End Function
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Error_Handler
Dim vCtrlDown, vShiftDown As Variant
Dim i As Integer
Dim bLoadupOne As Boolean
 
If g_iDebug = 13 Then
    Call InfoLog("frmPDD Form_KeyDown - Begin")
End If
 
  If KeyCode = oKeyDefs.Escape_Key Then
      bEscPressedOnError = True
  End If
 
  If gbLoadPreDispatched = True Or m_bInitializing Then
      KeyCode = 0
      If g_iDebug = 13 Then
         Call InfoLog("frmPDD Form_KeyDown Loadpredispatched = true and initializing - End")
      End If
      Exit Sub
  End If
   
  oKeyDefs.CheckForEnterKey Me.hwnd
  
  If Me.ActiveControl Is txtDolly(iDollyIndex) And Not IsUserShiftTabbing And KeyCode <> oKeyDefs.Keypad_Execute_Key Then
     If Len(ctlBayNr.Item(iBayIndex).Text) = 0 And Not bPredispatchEmpty(iBayIndex) And grdBayDetails(iBayIndex).Columns(GRD1_COL_PD).Value = HFCS_INDICATOR_FALSE Then
        KeyCode = 0
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD Form_KeyDown no bay number - End")
        End If
        Exit Sub
     End If
  End If
 
  If Me.ActiveControl Is cboDow Then
     If KeyCode = oKeyDefs.Escape_Key Then
        bExitScreen = True
        Unload Me
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD Form_KeyDown Escape Key - End")
        End If
        Exit Sub
     ElseIf KeyCode = oKeyDefs.Enter_Key Then
        EnterAsTab KeyCode
     End If
  End If
 
  If Me.ActiveControl Is grdMultLegView Then
    If KeyCode = oKeyDefs.Enter_Key Then
        EnterAsTab KeyCode
    End If
  End If
 
  Select Case KeyCode
    Case oKeyDefs.Escape_Key  'Escape pressed
        If (Me.ActiveControl Is ctlBayNr.Item(iBayIndex)) And bBayCompeleted(iBayIndex) Then
           Exit Sub
        ElseIf Me.Enabled = True Then
           Call ClearField
        End If
    
    
    Case oKeyDefs.F1_Key
         Form_Help Me
    Case oKeyDefs.Keypad_Execute_Key, vbKeyF10
       If bDisableF10 Then
          If g_iDebug = 13 Then
             Call InfoLog("frmPDD Form_KeyDown disable F10 - End")
          End If
          Exit Sub
          bDisableF10 = False
       Else
          DoSaveData
       End If
    Case vbKeyDelete
  
End Select
If g_iDebug = 13 Then
    Call InfoLog("frmPDD Form_KeyDown - End")
End If
Exit Sub
Error_Handler:
  glErrNum = 400
  gsError = "Form_KeyDown"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
End If
End Sub
 
Public Function DeleteLoadOnScreen(iRow As Integer, bAll As Boolean)
   On Error GoTo ErrorHandler
   Dim i As Integer
   Dim sBayNr As String
   Dim sTrailerNr As String
   Dim j As Integer
  
 If g_iDebug = 13 Then
    Call InfoLog("frmPDD DeleteLoadOnScreen - Begin")
End If
  
   If bAll Then
      For i = pInst.PDInfo.hPDHll.Count To 1 Step -1
         Set oColLoads(i - 1) = Nothing
         For j = 0 To 2
            Set oOldColLoads(i - 1, j) = Nothing
         Next j
         pInst.PDInfo.hPDHll.Remove i
      Next i
      If g_colBays.Count > 0 Then
        For i = g_colBays.Count To 1 Step -1
          g_colBays.Remove (i)
        Next i
      End If
   Else
      If pInst.PDInfo.hPDHll.Count >= iRow + 1 Then
         pInst.PDInfo.hPDHll.Remove iRow + 1
       
      End If
      If g_colBays.Count > 0 And iRow < g_colBays.Count Then
          g_colBays.Remove iRow + 1
      End If
  
   End If
  
   DisplayLdData pInst.PDInfo.hPDHll
  
   If pInst.PDInfo.hPDHll.Count = 0 Then
       ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
       mnuF11CancelPredispatch.Enabled = False
   End If
  
   Call SetF2F7Button
  
   If Not pInst.bArriveFlag Then
       Me.Show
   End If
  
   For i = 1 To 3
      If raArrivalData(i).iSelected = iBayIndex Then
         raArrivalData(i).iSelected = -999
         Exit For
      End If
   Next i
     
      
   bBalNrValidated(iBayIndex) = False
   bDollyValidated(iBayIndex) = False
   bBayCompeleted(iBayIndex) = False
   bCreateOutboundRecord(iBayIndex) = False
   bPredispatchEmpty(iBayIndex) = False
   bPredispatchEmptyWithTrailer(iBayIndex) = False
  
   lOrgFRDKey(iBayIndex) = 0
   sPreBayNum(iBayIndex) = vbNull
  
   bLostFocus = False
   gbProcessRetData = False
  
   'Clear the warning of the control bay
   If ctlWarningPopup.WarningVisible Then
        ctlWarningPopup.ClearWarning
    End If
   
    If iBayIndex < 2 Then
        bDatafromMatchingScreen(iBayIndex) = bDatafromMatchingScreen(iBayIndex + 1)
        bDataFromEloc(iBayIndex) = bDataFromEloc(iBayIndex + 1)
        bFromGTForm(iBayIndex) = bFromGTForm(iBayIndex + 1)
    Else
        bDatafromMatchingScreen(iBayIndex) = False
        bFromGTForm(iBayIndex) = False
        bDataFromEloc(iBayIndex) = False
    End If
    ' TT000394
    'If (iBayIndex > 1) Then
    '    grdBayDetails(iBayIndex - 1).Enabled = True
    'End If
   
    sPreJobNr = ""
    iDollyIndex = 0
    bDuplicateLoad = False
    bFromF4F5 = False
    sServiceTypeCD = ""
    iMultiLds = 0
    mbtab = False
    iBayIndex = 0
    bReLoad = False
    bUType = False
   
    bMultiForm = False
    bAtGTForm = False
    bMatchingForm = False
    iMultiLds = 0
    lblDeleteRecord.Visible = False
    bLostFocus = False
    bDisableF10 = False
    bOpenPreScreen = False
    bExitScreen = False
    bESC = False
    bSaveData = False
    bFinished = False
    iMyRc = 0
If g_iDebug = 13 Then
    Call InfoLog("frmPDD DeleteLoadOnScreen - End")
End If
    Exit Function
ErrorHandler:
    glErrNum = 400
   gsError = "DeleteLoadOnScreen"
   Call oProc.update_error_object(Me, gsError)
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
       ' Unload Me
    End If
 
End Function
 
 
Public Function ScreenCleanUp()
On Error GoTo Error_Handler
    If g_iDebug = 13 Then
       Call InfoLog("ScreenCleanUp - Begin")
    End If
   txtJobNr.Text = ""
   txtDriverName.Text = ""
   txtTractorNumber.Text = ""
   txtEloc.Text = ""
   Call xaMultiLegArray.ReDim(0, -1, 0, 9)
   grdMultLegView.Array = xaMultiLegArray
   grdMultLegView.ReBind
   chkDriverNotified.Value = 0
       
   If pInst.PDInfo.hPDHll.Count > 0 Then
      Call DeleteLoadOnScreen(0, True)
   End If
    If g_iDebug = 13 Then
       Call InfoLog("ScreenCleanUp - End")
    End If
    Exit Function
Error_Handler:
glErrNum = 400
gsError = "ScreenCleanUp"
'the object has disconnect from its clients grid error, retry the grid rebind
If Err.Number = -2147417848 Then
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField - received error -2147417848")
    End If
     
    If RecoverGridError(grdMultLegView) = True Then
        Resume Next
    End If
End If
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Function
 
'=============================================================================
' Description: set the forms tab order
'=============================================================================
Private Sub ReTab(lList As String)
 
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD - ReTab - Begin")
    End If
 
  cboDow.TabIndex = 0
  ctlWkendingDate.TabIndex = 1  'For tracker 1404.
 
  txtJobNr.TabIndex = 2                ' To prevent leaving control
  txtDriverName.TabIndex = 3
 
  grdMultLegView.TabIndex = 4
 
'  txtTractorNumber.TabIndex = 4
'  txtEloc.TabIndex = 5
  TextEnd.TabIndex = 5
 
  ctlBayNr(0).TabIndex = 6
'  If lList = LIST_GRIDS Then ' "LIST_Grids" Then
    grdBayDetails.Item(0).TabIndex = 7
'  Else
'    grdBayDetails.Item(0).TabIndex = 50
'  End If
 
  txtDolly(0).TabIndex = 8
  chkExtraDolly(0).TabIndex = 9
 
  ctlBayNr(1).TabIndex = 10
  ' TT000394
'  If lList = LIST_GRIDS Then ' "LIST_Grids" Then
    grdBayDetails.Item(1).TabIndex = 11
'  Else
'    grdBayDetails.Item(1).TabIndex = 55
'  End If
  txtDolly(1).TabIndex = 12
  chkExtraDolly(1).TabIndex = 13
 
  ctlBayNr(2).TabIndex = 14
  ' TT000394
'  If lList = LIST_GRIDS Then ' "LIST_Grids" Then
    grdBayDetails.Item(2).TabIndex = 15
'  Else
'    grdBayDetails.Item(2).TabIndex = 60
'  End If
  txtDolly(2).TabIndex = 16
'  chkDriverNotified.TabIndex = 13
    If g_iDebug = 13 Then
       Call InfoLog("frm PDD - ReTab - End")
    End If
    Exit Sub
Error_Handler:
glErrNum = 400
gsError = "ReTab"
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
        Resume Next
End If
End Sub
 
'===========================================================
' Sub:          repositionFormControls()
' Description:  when user turns on/off toolbar, resize and
'               reposition controls on the form
'===========================================================
Private Sub repositionFormControls(bIsTbarOn As Boolean)
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD repositionFormControls - Begin")
    End If
    Dim iHeight As Integer
    Dim TBARHEIGHT As Integer
   
    TBARHEIGHT = 240
   
    If bIsTbarOn = True Then
        iHeight = TBARHEIGHT
    Else
        iHeight = -TBARHEIGHT
    End If
   
    frmBayData.Top = frmBayData.Top + 3 * iHeight
    frmBayData.Height = frmBayData.Height
   
    fraJobInformation.Top = fraJobInformation.Top + 3 * iHeight
    fraJobInformation.Height = fraJobInformation.Height
   
    lblDeleteRecord.Top = lblDeleteRecord.Top + 3 * iHeight
    lblDeleteRecord.Height = lblDeleteRecord.Height
   
    'Adjust all the text fields
    cboDow.Top = cboDow.Top + 3 * iHeight
    ctlWkendingDate.Top = ctlWkendingDate.Top + 3 * iHeight
    txtJobNr.Top = txtJobNr.Top + 3 * iHeight
    txtDriverName.Top = txtDriverName.Top + 3 * iHeight
    txtTractorNumber.Top = txtTractorNumber.Top + 3 * iHeight
    txtEloc.Top = txtEloc.Top + 3 * iHeight
    chkDriverNotified.Top = chkDriverNotified.Top + 3 * iHeight
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD repositionFormControls - End")
    End If
End Sub
 
Public Sub ShiftLdInfo()
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ShiftLdInfo - Begin")
    End If
    pInst.PDInfo.hPDHll.Add pInst.PDInfo.hPDHll.Item(iBayIndex + 1)
    pInst.PDInfo.hPDHll.Remove iBayIndex + 1
    bCreateOutboundRecord(iBayIndex) = False
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ShiftLdInfo - End")
    End If
End Sub
 
Private Function ClearField()
On Error GoTo Error_Handler
   Dim i As Integer
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD ClearField - Begin")
   End If
   mbClearData = True
'  bGridGotFocus = False
'  bRoutingCanceled = False
   If ActiveControl.Name = "cboDow" And Not bExitScreen Then
     mbClearData = False
     bExitScreen = True
     If g_iDebug = 13 Then
       Call InfoLog("frmPDD ClearField cboDow - End")
     End If
     Unload Me
     Exit Function
   End If
  
   If ActiveControl.Name = "ctlWkendingDate" And Not bExitScreen Then
     mbClearData = False
     bExitScreen = True
     If g_iDebug = 13 Then
       Call InfoLog("frmPDD ClearField ctlWkendingDate - End")
     End If
     Unload Me
     Exit Function
   End If
  
   If ActiveControl.Name = "txtJobNr" And Not bExitScreen Then
      If txtJobNr.Text <> "" Then
         txtJobNr.Text = ""
         sPreJobNr = ""
         If Not (xaMultiLegArray Is Nothing) Then
            If (xaMultiLegArray.UpperBound(1) <> -1) Then
                grdMultLegView.Bookmark = 0
'               If (xaMultiLegArray(xaMultiLegArray.UpperBound(1), 8) = True) Then
'                   Call ResetForm
'               End If
            End If
         End If
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtJobNr not blank - End")
         End If
         mbClearData = False
         Exit Function
      Else
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtJobNr blank - End")
         End If
         Call ResetForm
         bExitScreen = True
         Exit Function
      End If
   ElseIf ActiveControl.Name = "txtJobNr" And bExitScreen Then
      mbClearData = False
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField txtJobNr active - End")
     End If
      Unload Me
      Exit Function
   End If
  
   If ActiveControl.Name = "txtDriverName" Then
      If txtDriverName <> "" Then
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtDriverName not blank - End")
         End If
         mbClearData = False
         txtDriverName = ""
         If (bFromGtF8 Or pInst.bArriveFlag) And xaMultiLegArray.UpperBound(1) = -1 Then
            bExitScreen = True
         End If
         Exit Function
      Else
         If bExitScreen = True And (bFromGtF8 Or pInst.bArriveFlag) Then
            If g_iDebug = 13 Then
                Call InfoLog("frmPDD ClearField txtDriverName - End - If bExitScreen = True And...")
            End If
            Unload Me
            Exit Function
         End If
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtDriverName blank - End")
         End If
         Call ResetForm
         bExitScreen = True
         Exit Function
      End If
   ElseIf bExitScreen Then
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField txtDriverName bExitScreen - End")
      End If
      mbClearData = False
      Unload Me
      Exit Function
   End If
     
   If ActiveControl.Name = "grdMultLegView" Then
      If xaMultiLegArray.UpperBound(1) <> -1 Then
'        Call xaMultiLegArray.ReDim(0, -1, 0, 9)
'        grdMultLegView.Array = xaMultiLegArray
'        grdMultLegView.ReBind
         If txtDriverName.Enabled = False Then txtDriverName.Enabled = True
         txtDriverName.SetFocus
      Else
        mbClearData = False
        Call ResetForm
        bExitScreen = True
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField - End - If ActiveControl.Name = 'grdMultLegView' 1")
        End If
        Exit Function
      End If
  ElseIf bExitScreen Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField - End - If bExitScreen")
        End If
      mbClearData = False
      Unload Me
      Exit Function
   End If
     
   If ActiveControl.Name = "txtTractorNumber" Then
      If txtTractorNumber <> "" Then
         mbClearData = False
         txtTractorNumber = ""
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtTractorNumber not blank - End")
         End If
         If (pInst.bArriveFlag Or bFromGtF8) And Len(txtEloc.Text) = 0 Then
            bExitScreen = True
         End If
         Exit Function
      Else
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtTractorNumber blank - End")
         End If
         Call ResetForm
         bExitScreen = True
        
         If Not pInst.bArriveFlag And Not bFromGtF8 Or Len(txtEloc.Text) = 0 Then
            pInst.bArriveFlag = False
         End If
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField txtTractorNumber exit screen - End")
      End If
      Unload Me
      Exit Function
   End If
  
   If ActiveControl.Name = "txtEloc" Then
      If txtEloc <> "" Then
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtEloc not blank - End")
         End If
         mbClearData = False
         txtEloc = ""
         Exit Function
      Else
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField txtEloc blank - End")
         End If
         Call ResetForm
         bExitScreen = True
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField txtEloc ExitScreen - End")
      End If
      Unload Me
      Exit Function
   End If
  
   ' TT000394
   If (ActiveControl.Name = "ctlBayNr" Or ActiveControl.Name = "grdBayDetails") And iBayIndex = 0 Then
      'TT000394
      If ActiveControl.Name = "grdBayDetails" Then
         SetGridNormal (iBayIndex)
         ctlBayNr(iBayIndex).Visible = True
         ctlBayNr(iBayIndex).Enabled = True
         bBayCompeleted(iBayIndex) = False
         ctlBayNr(iBayIndex).SetFocus
      End If
      If Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And bBayCompeleted(iBayIndex) Then
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank - bay completed - End")
         End If
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count = iBayIndex + 1 And Not bNeedClearLds Then
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank - count is bay index - End")
         End If
         ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
         For i = 1 To 3
           If Format$(pInst.raLockedBay(i), "0000") = ctlBayNr.Item(iBayIndex).Text Then
               pInst.raLockedBay(i) = 0
               Exit For
           End If
         Next i
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 0
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField If (ActiveControl.Name = 'ctlBayNr' Or - End")
         End If
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And Not bNeedClearLds Then
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank - not need clear loads - End")
         End If
         ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
         For i = 1 To 3
           If Format$(pInst.raLockedBay(i), "0000") = ctlBayNr.Item(iBayIndex).Text Then
               pInst.raLockedBay(i) = 0
               Exit For
           End If
         Next i
        
         Call DeleteLoadOnScreen(iBayIndex, False)
         If pInst.PDInfo.hPDHll.Count = 0 Then
            grdBayDetails(iBayIndex).ReBind
            bNeedClearLds = False
         End If
         iBayIndex = 0
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And bNeedClearLds Then
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank - need clear loads - End")
         End If
         ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
         For i = 1 To 3
           If Format$(pInst.raLockedBay(i), "0000") = ctlBayNr.Item(iBayIndex).Text Then
               pInst.raLockedBay(i) = 0
               Exit For
           End If
         Next i
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 0
         Exit Function
      ElseIf IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count >= iBayIndex + 1 Then
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr is blank - count is bay index - End")
         End If
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 0
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And bBayCompeleted(iBayIndex) Then
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank - bay completed - End")
         End If
         Exit Function
      Else
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr reset form - End")
         End If
         Call ResetForm
         bExitScreen = True
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField ctlBayNr exit screen - End")
     End If
      Unload Me
      Exit Function
   End If
  
   ' TT000394
   If (ActiveControl.Name = "ctlBayNr" Or ActiveControl.Name = "grdBayDetails") And iBayIndex = 1 Then
      'TT000394
      If ActiveControl.Name = "grdBayDetails" Then
         SetGridNormal (iBayIndex)
         ctlBayNr(iBayIndex).Visible = True
         ctlBayNr(iBayIndex).Enabled = True
         bBayCompeleted(iBayIndex) = False
         ctlBayNr(iBayIndex).SetFocus
     End If
      If Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And bBayCompeleted(iBayIndex) Then
         If bBalNrValidated(iBayIndex) Then
            mbClearData = False
            If g_iDebug = 13 Then
              Call InfoLog("frmPDD ClearField ctlBayNr bBalNrValidate - End")
            End If
            Exit Function
         Else
            mbClearData = False
            ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
            ctlBayNr.Item(iBayIndex).Text = ""
            iBayIndex = 1
            If g_iDebug = 13 Then
              Call InfoLog("frmPDD ClearField ctlBayNr not bBalNrValidate - End")
            End If
            Exit Function
         End If
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count = iBayIndex + 1 Then
         mbClearData = False
         ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 1
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank bay - bay index 1 - count is bay index - End")
         End If
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) Then
         mbClearData = False
         ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 1
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank bay - bay index 1 - End")
         End If
         Exit Function
      ElseIf IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count >= iBayIndex + 1 Then
         mbClearData = False
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 1
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank bay - bay index 1 - count > bay index - End")
         End If
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count > iBayIndex + 1 Then
         mbClearData = False
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr not blank bay - bay index 1 - count > bay index  - End")
         End If
         Exit Function
      Else
         mbClearData = False
         Call ResetForm
         bExitScreen = True
         If g_iDebug = 13 Then
           Call InfoLog("frmPDD ClearField ctlBayNr reset form - bay index 1 - End")
         End If
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      Unload Me
      If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField ctlBayNr exit screen - bay index 1 - End")
      End If
      Exit Function
   End If
  
   ' TT000394
   If (ActiveControl.Name = "ctlBayNr" Or ActiveControl.Name = "grdBayDetails") And iBayIndex = 2 Then
      'TT000394
      If ActiveControl.Name = "grdBayDetails" Then
         SetGridNormal (iBayIndex)
         ctlBayNr(iBayIndex).Visible = True
         ctlBayNr(iBayIndex).Enabled = True
         bBayCompeleted(iBayIndex) = False
         ctlBayNr(iBayIndex).SetFocus
      End If
      If Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count > iBayIndex + 1 Then
         mbClearData = False
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField ctlBayNr - bay index 2 not blank count > bayindex - End")
         End If
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count = iBayIndex + 1 Then
         mbClearData = False
         ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 2
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField ctlBayNr - bay index 2 not blank  - count bay index - End")
         End If
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) Then
         mbClearData = False
         ProcessLockBay ctlBayNr.Item(iBayIndex).Text, Nothing, BAY_UNLOCK
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 2
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField ctlBayNr - bay index 2 not blank - End")
         End If
         Exit Function
      ElseIf Not IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count > iBayIndex + 1 Then
         mbClearData = False
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField ctlBayNr - bay index 2 not blank - count > bayindex - End")
         End If
         Exit Function
      ElseIf IsBlank(ctlBayNr.Item(iBayIndex).Text) And pInst.PDInfo.hPDHll.Count >= iBayIndex + 1 Then
         mbClearData = False
         Call DeleteLoadOnScreen(iBayIndex, False)
         ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
         mnuF11CancelPredispatch.Enabled = False
         iBayIndex = 2
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField ctlBayNr - bay index 2 blank - count > bayindex - End")
         End If
         Exit Function
      Else
         mbClearData = False
         Call ResetForm
         bExitScreen = True
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField ctlBayNr - bay index 2 reset form - End")
         End If
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      Unload Me
      If g_iDebug = 13 Then
         Call InfoLog("frmPDD ClearField ctlBayNr - bay index 2 - exit screen  - End")
      End If
      Exit Function
   End If
  
   If (ActiveControl.Name = "txtDolly" Or ActiveControl.Name = "chkExtraDolly") And iDollyIndex = 0 Then
      If Not IsBlank(txtDolly.Item(iDollyIndex).Text) Then
         mbClearData = False
         txtDolly.Item(iDollyIndex).Text = ""
         If iDollyIndex <> 2 Then
            chkExtraDolly(iDollyIndex).Value = 0
            chkExtraDolly(iDollyIndex).Enabled = False
           chkExtraDolly(iDollyIndex).TabStop = False
         End If
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField txtDolly not blank - dolly 0 - End")
         End If
         Exit Function
       Else
         mbClearData = False
         Call ResetForm
         bExitScreen = True
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField txtDolly blank - dolly 0 - End")
         End If
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      Unload Me
      If g_iDebug = 13 Then
         Call InfoLog("frmPDD ClearField txtDolly not blank - dolly 0 - exit screen  - End")
      End If
      Exit Function
   End If
  
   If (ActiveControl.Name = "txtDolly" Or ActiveControl.Name = "chkExtraDolly") And iDollyIndex = 1 Then
      If Not IsBlank(txtDolly.Item(iDollyIndex).Text) Then
         mbClearData = False
         txtDolly.Item(iDollyIndex).Text = ""
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField txtDolly not blank - dolly 1 - End")
         End If
         Exit Function
       Else
         mbClearData = False
         Call ResetForm
         bExitScreen = True
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField txtDolly blank - dolly 1 - End")
         End If
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      Unload Me
      If g_iDebug = 13 Then
         Call InfoLog("frmPDD ClearField txtDolly not blank - dolly 1 - exit screen - End")
      End If
      Exit Function
   End If
  
   If (ActiveControl.Name = "txtDolly" Or ActiveControl.Name = "chkExtraDolly") And iDollyIndex = 2 Then
      If Not IsBlank(txtDolly.Item(iDollyIndex).Text) Then
         mbClearData = False
         txtDolly.Item(iDollyIndex).Text = ""
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField txtDolly not blank - dolly 2 - End")
         End If
         Exit Function
       Else
         mbClearData = False
         Call ResetForm
         bExitScreen = True
         If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField txtDolly blank - dolly 2 - End")
         End If
         Exit Function
      End If
   ElseIf bExitScreen Then
      mbClearData = False
      Unload Me
      If g_iDebug = 13 Then
         Call InfoLog("frmPDD ClearField txtDolly not blank - dolly 2 - exit screen - End")
      End If
      Exit Function
   End If
   If g_iDebug = 13 Then
      Call InfoLog("frmPDD ClearField - End")
   End If
   mbClearData = False
   Exit Function
Error_Handler:
glErrNum = 400
gsError = "ClearField"
'the object has disconnect from its clients grid error, retry the grid rebind
If Err.Number = -2147417848 Then
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField - received error -2147417848")
    End If
     
    If RecoverGridError(grdBayDetails(iBayIndex)) = True Then
        Resume Next
    End If
End If
Call oProc.update_error_object(Me, gsError)
 
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
  End If
End Function
 
Private Function ResetForm()
On Error GoTo Error_Handler
 
Dim i As Integer
Dim j As Integer
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ResetForm - Begin")
    End If
 
        If ctlWarningPopup.WarningVisible Then
           ctlWarningPopup.ClearWarning
        End If
       
        iPredispatch = 0
        pInst.icdl = 0
        m_bSettingLoadInfo = False
'        pInst.bArriveFlag = False
        If Not pInst.bArriveFlag And Not bFromGtF8 Then
            txtJobNr.Text = ""
            txtDriverName.Text = ""
        End If
       
        txtTractorNumber.Text = ""
        txtEloc.Text = ""
        chkDriverNotified.Value = 0
        cboDow.Enabled = True
        ctlWkendingDate.Enabled = True
       
        If Not xaMultiLegArray Is Nothing Then
            Call xaMultiLegArray.ReDim(0, -1, 0, 9)
            grdMultLegView.Array = xaMultiLegArray
            grdMultLegView.ReBind
        End If
       
        iMyRc = UMBID_DELETE
        If pInst.PDInfo.hPDHll.Count > 0 Then
           Call DeleteLoadOnScreen(0, True)
        End If
          
        For i = 1 To 3
           ClearLoadDisplay i - 1
'           chkCancelPD(i - 1).Value = 0
'           chkCancelPD(i - 1).Enabled = False
           Set oColLoads(i - 1) = Nothing
           For j = 0 To 2
            Set oOldColLoads(i - 1, j) = Nothing
           Next j
        Next i
       
        ctlToolbarManager.buttonEnabled BT_F11_CANCEL_PD, False
        ctlToolbarManager.buttonEnabled BT_F6_UPD_LEG, False
       
        If Not pInst.bArriveFlag And Not bFromGtF8 Then
            szFdrSchDow = vbNullString
            szFdrSchWndDT = vbNullString
        End If
       
        bFromGTForm(0) = False
        bFromGTForm(1) = False
        bFromGTForm(2) = False
        bDataFromEloc(0) = False
        bDataFromEloc(1) = False
        bDataFromEloc(2) = False
        bDatafromMatchingScreen(0) = False
        bDatafromMatchingScreen(1) = False
        bDatafromMatchingScreen(2) = False
         
        sPreJobNr = ""
        bShiftLdInfo = False
        iDollyIndex = 0
        bFromF4F5 = False
        sServiceTypeCD = ""
        iMultiLds = 0
        mbtab = False
        iBayIndex = 0
        bReLoad = False
        bUType = False
       
        lOrgFRDKey(0) = 0
        lOrgFRDKey(1) = 0
        lOrgFRDKey(2) = 0
       
        bBalNrValidated(0) = False
        bBalNrValidated(1) = False
        bBalNrValidated(2) = False
        bBayCompeleted(0) = False
        bBayCompeleted(1) = False
        bBayCompeleted(2) = False
        bDollyValidated(0) = False
        bDollyValidated(1) = False
        bDollyValidated(2) = False
        bctlBayValidateComplete(0) = False
        bctlBayValidateComplete(1) = False
        bctlBayValidateComplete(2) = False
       
        bCreateOutboundRecord(0) = False
        bCreateOutboundRecord(1) = False
        bCreateOutboundRecord(2) = False
        sPreBayNum(0) = ""
        sPreBayNum(1) = ""
        sPreBayNum(2) = ""
   
        bMultiForm = False
        bAtGTForm = False
        bMatchingForm = False
       
        bMatchingLoadFound(0) = False
        bMatchingLoadFound(1) = False
        bMatchingLoadFound(2) = False
               
        iMultiLds = 0
        lblDeleteRecord.Visible = False
        bLostFocus = False
        bDisableF10 = False
        bOpenPreScreen = False
        bExitScreen = False
        bESC = False
        bSaveData = False
        bFinished = False
     
    
        iMyRc = 0
        bESC = True
        bExitScreen = True
       
        If Not pInst.bArriveFlag And Not bFromGtF8 Then
            txtJobNr.Enabled = True
            txtJobNr.SetFocus 'QS 6/24
            txtLastDriverName = vbNullString
       Else
        '    txtEloc.Enabled = True
            grdMultLegView.Enabled = True
            txtDriverName.Enabled = True
            txtDriverName.SetFocus
        '    txtTractorNumber.Enabled = True
        '    txtTractorNumber.SetFocus
        End If
       
        bDuplicateLoad = False
        ' TT000394
        ctlBayNr.Item(0).Enabled = True
        ctlBayNr.Item(1).Enabled = True
        ctlBayNr.Item(2).Enabled = True
 
        txtDolly.Item(0).Enabled = False
        txtDolly.Item(1).Enabled = False
        txtDolly.Item(2).Enabled = False
       
        ' TT000394
        bPredispatchEmpty(0) = False
        bPredispatchEmpty(1) = False
        bPredispatchEmpty(2) = False
   
        bPredispatchEmptyWithTrailer(0) = False
        bPredispatchEmptyWithTrailer(1) = False
        bPredispatchEmptyWithTrailer(2) = False
       
        Call SetColumnLocksForEmptyLoad(0, False)
       
        grdBayDetails(0).Enabled = True
        grdBayDetails(1).Enabled = True
        grdBayDetails(2).Enabled = True
 
        Call ReTab(LIST_NO_GRID)
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ResetForm - End")
    End If
    Exit Function
Error_Handler:
glErrNum = 400
gsError = "ResetForm"
'the object has disconnect from its clients grid error, retry the grid rebind
If Err.Number = -2147417848 Then
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD ClearField - received error -2147417848")
    End If
     
    If RecoverGridError(grdMultLegView) = True Then
        Resume Next
    End If
End If
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
   '     Unload Me
End If
End Function
 
Public Function IsUserShiftTabbing() As Boolean
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD IsUserShiftTabbing - Begin")
    End If
   
IsUserShiftTabbing = False
  If GetKeyState(vbKeyTab) < 0 And GetKeyState(vbKeyShift) < 0 Then
    IsUserShiftTabbing = True
  End If
  If g_iDebug = 13 Then
       Call InfoLog("frmPDD IsUserShiftTabbing - End")
    End If
    Exit Function
Error_Handler:
glErrNum = 400
gsError = "IsUserShiftTabbing"
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
End If
End Function
 
Private Function CheckDuplicate(Index As Integer) As Boolean
On Error GoTo Error_Handler
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CheckDuplicate - Begin - Index: " & Index)
    End If
 
CheckDuplicate = False
If bPredispatchEmpty(Index) = True Or (Not IsNull(pInst.PDInfo.hPDHll(Index + 1).AssignedPDOID) And Len(ctlBayNr(Index).Text) = 0) Then
    bDuplicateLoad = False
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CheckDuplicate bPredispatchEmpty - End - Index: " & Index)
    End If
    Exit Function
End If
 
If Index + 1 = 2 Then
        If (pInst.PDInfo.hPDHll.Item(1).szTlrNr = pInst.PDInfo.hPDHll.Item(2).szTlrNr And bPredispatchEmpty(Index) = False) Then ' And _
'          ((StrComp(pInst.PDInfo.hPDHll.Item(1).szOrigin, pInst.PDInfo.hPDHll.Item(2).szOrigin)) = 0 And _
'              (StrComp(pInst.PDInfo.hPDHll.Item(1).szOriginCn, pInst.PDInfo.hPDHll.Item(2).szOriginCn)) = 0 And _
'                (StrComp(pInst.PDInfo.hPDHll.Item(1).szDestin, pInst.PDInfo.hPDHll.Item(2).szDestin)) = 0 And _
'                     (StrComp(pInst.PDInfo.hPDHll.Item(1).szDestinCn, pInst.PDInfo.hPDHll.Item(2).szDestinCn)) = 0 And _
'                        (StrComp(pInst.PDInfo.hPDHll.Item(1).szDestinSrt, pInst.PDInfo.hPDHll.Item(2).szDestinSrt)) = 0) And _
'                          (pInst.PDInfo.hPDHll.Item(1).iSequenceNr = pInst.PDInfo.hPDHll.Item(2).iSequenceNr)
                           ctlWarningPopup.ShowWarning ctlBayNr(1).hwnd, LoadResString(sDuplicateload), 2000
                           If pInst.PDInfo.hPDHll.Count = Index + 1 Then
                              pInst.PDInfo.hPDHll.Remove pInst.PDInfo.hPDHll.Count  'Remove the duplicate one  QS 8/2/05
                           End If
                           bBalNrValidated(Index) = False
                           ctlBayNr.Item(1).SelStart = 0
                           ctlBayNr.Item(1).SelLength = Len(ctlBayNr.Item(Index).Text)
                           CheckDuplicate = True
                           bDuplicateLoad = True
                           bReLoad = False
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CheckDuplicate end index = 2 - End - Index: " & Index)
    End If
                          
                           Exit Function
    Else
       bDuplicateLoad = False
    End If
   End If
    If Index + 1 = 3 Then
        If (pInst.PDInfo.hPDHll.Item(1).szTlrNr = pInst.PDInfo.hPDHll.Item(3).szTlrNr And bPredispatchEmpty(Index) = False) Then ' And _
'           ((StrComp(pInst.PDInfo.hPDHll.Item(1).szOrigin, pInst.PDInfo.hPDHll.Item(3).szOrigin)) = 0 And _
'              (StrComp(pInst.PDInfo.hPDHll.Item(1).szOriginCn, pInst.PDInfo.hPDHll.Item(3).szOriginCn)) = 0 And _
'                (StrComp(pInst.PDInfo.hPDHll.Item(1).szDestin, pInst.PDInfo.hPDHll.Item(3).szDestin)) = 0 And _
'                     (StrComp(pInst.PDInfo.hPDHll.Item(1).szDestinCn, pInst.PDInfo.hPDHll.Item(3).szDestinCn)) = 0 And _
'                        (StrComp(pInst.PDInfo.hPDHll.Item(1).szDestinSrt, pInst.PDInfo.hPDHll.Item(3).szDestinSrt)) = 0) And _
'                          (pInst.PDInfo.hPDHll.Item(1).iSequenceNr = pInst.PDInfo.hPDHll.Item(3).iSequenceNr) Then
                            ctlWarningPopup.ShowWarning ctlBayNr(2).hwnd, LoadResString(sDuplicateload), 2000
                            If pInst.PDInfo.hPDHll.Count = Index + 1 Then
                               pInst.PDInfo.hPDHll.Remove pInst.PDInfo.hPDHll.Count  'Remove the duplicate one  QS 8/2/05
                            End If
                            bBalNrValidated(Index) = False
                            ctlBayNr.Item(2).SelStart = 0
                            ctlBayNr.Item(2).SelLength = Len(ctlBayNr.Item(Index).Text)
                            CheckDuplicate = True
                            bDuplicateLoad = True
                            bReLoad = False
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CheckDuplicate end = 3 - End - Index: " & Index)
    End If
                           
                            Exit Function
                           
       Else
          bDuplicateLoad = False
       End If
         If (pInst.PDInfo.hPDHll.Item(2).szTlrNr = pInst.PDInfo.hPDHll.Item(3).szTlrNr And bPredispatchEmpty(Index) = False) Then ' And _
'            ((StrComp(pInst.PDInfo.hPDHll.Item(2).szOrigin, pInst.PDInfo.hPDHll.Item(3).szOrigin)) = 0 And _
'              (StrComp(pInst.PDInfo.hPDHll.Item(2).szOriginCn, pInst.PDInfo.hPDHll.Item(3).szOriginCn)) = 0 And _
'                (StrComp(pInst.PDInfo.hPDHll.Item(2).szDestin, pInst.PDInfo.hPDHll.Item(3).szDestin)) = 0 And _
'                     (StrComp(pInst.PDInfo.hPDHll.Item(2).szDestinCn, pInst.PDInfo.hPDHll.Item(3).szDestinCn)) = 0 And _
'                        (StrComp(pInst.PDInfo.hPDHll.Item(2).szDestinSrt, pInst.PDInfo.hPDHll.Item(3).szDestinSrt)) = 0) And _
'                          (pInst.PDInfo.hPDHll.Item(2).iSequenceNr = pInst.PDInfo.hPDHll.Item(3).iSequenceNr) Then
                         ctlWarningPopup.ShowWarning ctlBayNr(2).hwnd, LoadResString(sDuplicateload), 2000
                         If pInst.PDInfo.hPDHll.Count = Index + 1 Then
                            pInst.PDInfo.hPDHll.Remove pInst.PDInfo.hPDHll.Count  'Remove the duplicate one  QS 8/2/05
                         End If
                         bBalNrValidated(Index) = False
                         ctlBayNr.Item(2).SelStart = 0
                         ctlBayNr.Item(2).SelLength = Len(ctlBayNr.Item(Index).Text)
                         CheckDuplicate = True
                         bDuplicateLoad = True
                         bReLoad = False
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CheckDuplicate end bReload = false - END - Index: " & Index)
    End If
                        
                         Exit Function
                     
 
       Else
          bDuplicateLoad = False
       End If
 
    End If
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD CheckDuplicate - End - Index: " & Index)
    End If
    Exit Function
Error_Handler:
glErrNum = 400
gsError = "CheckDuplicate"
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
End If
End Function
 
Private Function InitialVaribles()
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD InitialVariables - Begin")
    End If
   
    bFromGTForm(0) = False
    bFromGTForm(1) = False
    bFromGTForm(2) = False
   
    bDataFromEloc(0) = False
    bDataFromEloc(1) = False
    bDataFromEloc(2) = False
   
    bDatafromMatchingScreen(0) = False
    bDatafromMatchingScreen(1) = False
    bDatafromMatchingScreen(2) = False
   
    iPredispatch = 0
    szFdrSchDow = vbNullString
    szFdrSchWndDT = vbNullString
    bNeedClearLds = False
    txtDriverName.Enabled = False
    txtTractorNumber.Enabled = False
    txtEloc.Enabled = False
    grdMultLegView.Enabled = False
    chkExtraDolly(0).Enabled = False
    chkExtraDolly(1).Enabled = False
    txtDolly.Item(0).Enabled = False
    txtDolly.Item(1).Enabled = False
    txtDolly.Item(2).Enabled = False
    chkDriverNotified.Enabled = False
      
    sPreJobNr = ""
    bMatchingLoadFound(0) = False
    bMatchingLoadFound(1) = False
    bMatchingLoadFound(2) = False
    bShiftLdInfo = False
   bIntxtDriverNa = False
    iDollyIndex = 0
    bDuplicateLoad = False
    bFromF4F5 = False
    sServiceTypeCD = ""
    iMultiLds = 0
    mbtab = False
    iBayIndex = 0
    bReLoad = False
    bUType = False
   
    lOrgFRDKey(0) = 0
    lOrgFRDKey(1) = 0
    lOrgFRDKey(2) = 0
    bBalNrValidated(0) = False
    bBalNrValidated(1) = False
    bBalNrValidated(2) = False
    bBayCompeleted(0) = False
    bBayCompeleted(1) = False
    bBayCompeleted(2) = False
    bctlBayValidateComplete(0) = False
    bctlBayValidateComplete(1) = False
    bctlBayValidateComplete(2) = False
    bDollyValidated(0) = False
    bDollyValidated(1) = False
    bDollyValidated(2) = False
    bCreateOutboundRecord(0) = False
    bCreateOutboundRecord(1) = False
    bCreateOutboundRecord(2) = False
    sPreBayNum(0) = ""
    sPreBayNum(1) = ""
    sPreBayNum(2) = ""
   
    bMultiBay(0) = False
    bMultiBay(1) = False
    bMultiBay(2) = False
   
    bMultiForm = False
    bAtGTForm = False
    bMatchingForm = False
    iMultiLds = 0
    lblDeleteRecord.Visible = False
    bLostFocus = False
    bDisableF10 = False
    bOpenPreScreen = False
    bExitScreen = False
    bESC = False
   
    bSaveData = False
    bFinished = False
    iMyRc = 0
  
    grdBayDetails(0).Bookmark = Null
    grdBayDetails(1).Bookmark = Null
    grdBayDetails(2).Bookmark = Null
    grdBayDetails(0).ReBind
    grdBayDetails(1).ReBind
    grdBayDetails(2).ReBind
   
    gbLoadPreDispatched = False
    g_bDriverPredispatched = False
    g_bJobPredispatched = False
    g_bELOCDiff = False
   
    mbClearData = False
   
    Set oYard = New HFCSYardObject.Yard
    Set oBay = New HFCSYardObject.Bay
   
    Set oRegistry = New HFRegistryObject.clsRegistry
    Set oWSInfo = New HFCSWSInformation.clsComm
   
    Set oWSNotify = New HFCSWSInformation.clsWSInfo
    Set OGlobalFormNames = New GLOBALDEFS.clsFormNames
    Set oGlobalMsgs = New GLOBALDEFS.clsMessages
    Set oEventlog = New HFErrorObject.clsEventLogs
    Set oColHeads = New GLOBALDEFS.clsColumnHeaders
    Set oHelp = New GLOBALDEFS.clsHelp
    Set oErrorObject = New HFErrorObject.clsErrorHandling
    Set oKeyDefs = New GLOBALDEFS.clsKeyCodes
    Set oFdrFns = New HFCSFeederGlobal.clsFeederGlobal
    Set oNotify = oWSNotify.NotificationClass
   
    Set oshell = oWSNotify.NotificationClass    'QS 11/16/04
    Set g_colBays = New Collection
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD InitialVariables - End")
    End If
    Exit Function
   
Error_Handler:
glErrNum = 400
gsError = "InitialVaribles"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                           IIf(Err.Number <> 0, Err.Number, glErrNum), _
                           Err.Description & " Module:" & gsError, _
                           oProc, _
                           ERROR_MSG, _
                           FEEDER_DISPATCH_DRIVER) Then
Set oEventlog = Nothing
'Unload Me
End If
End Function
 
Public Function SetF2F7Button()
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD SetF2F7Button - Begin")
    End If
 
If Len(ctlBayNr.Item(iBayIndex).Text) > 0 And pInst.PDInfo.hPDHll.Count > 0 Then  '5/24
 
            If pInst.PDInfo.hPDHll.Item(pInst.PDInfo.hPDHll.Count).szMultLdIr = IR_TRUE And (pInst.PDInfo.hPDHll.Item(pInst.PDInfo.hPDHll.Count).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(pInst.PDInfo.hPDHll.Count).lLdMsgId > 0) _
               And pInst.PDInfo.hPDHll.Item(pInst.PDInfo.hPDHll.Count).szMultLdIr = IR_TRUE Then
               SetAvailMenuAndTool mat_bay, True, True
            ElseIf pInst.PDInfo.hPDHll.Item(pInst.PDInfo.hPDHll.Count).lTlrMsgId > 0 Or pInst.PDInfo.hPDHll.Item(pInst.PDInfo.hPDHll.Count).lLdMsgId > 0 Then
               SetAvailMenuAndTool mat_bay, True, False
            Else
               SetAvailMenuAndTool mat_bay, True, False
            End If
        Else
           SetAvailMenuAndTool mat_bay, False, False
        End If
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD SetF2F7Button - End")
        End If
        Exit Function
       
Error_Handler:
glErrNum = 400
gsError = "SetF2F7Button"
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
End If
End Function
 
Public Function CheckDuplicateLoads2(oLoad As PdStruct) As Boolean
On Error GoTo Error_Handler
Dim i As Integer
 
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD CheckDuplicateLoads2 - Begin")
    End If
 
   CheckDuplicateLoads2 = False
  
       ' Check if the records match with respect to the search criteria
      If (StrComp(pInst.PDInfo.hPDHll.Item(i).szOrigin, oLoad.szOrigin)) = 0 Then
         If (StrComp(pInst.PDInfo.hPDHll.Item(i).szOriginCn, oLoad.szOriginCn)) = 0 Then
            If (StrComp(pInst.PDInfo.hPDHll.Item(i).szDestin, oLoad.szDestin)) = 0 Then
               If (StrComp(pInst.PDInfo.hPDHll.Item(i).szDestinCn, oLoad.szDestinCn)) = 0 Then
                  If (StrComp(pInst.PDInfo.hPDHll.Item(i).szDestinSrt, oLoad.szDestinSrt)) = 0 Then
                     CheckDuplicateLoads2 = True 'Duplicate load found
                    
                  End If
               End If
            End If
         End If
      End If
      If g_iDebug = 13 Then
            Call InfoLog("frmPDD CheckDuplicateLoads2 - End")
      End If
      Exit Function
Error_Handler:
  glErrNum = 400
  gsError = "CheckDuplicateLoads2"
  Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Function
 
Public Function AddPredispatchDriverData(lTransaction As Long, icount As Integer, _
                                        Optional UpdDriverNotifyOnly As Boolean = False) As Boolean
On Error GoTo ErrorHandler
Dim bResult As Boolean
Dim sFnName As String
Dim sSQL As String
Dim lCurForTlrEntity As Long
Dim sSequenceNr As String
Dim bInsert As Boolean
Dim sTableNa As String
Dim sLocHub As String
Dim rsRecord As ADODB.Recordset
Dim bSameFrdKey As Boolean
 
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD AddPredispatchDriverData - Begin")
End If
 
sFnName = "SavePredispatchDriverData"
 
sLocHub = oWSNotify.LocalHub
TempSlic = Right$(sLocHub, 5)
TempCny = Left$(sLocHub, 2)
bSameFrdKey = False
If bPredispatchEmptyWithTrailer(icount - 1) = True Then
    oLoad.PredispatchSequenceEdited = True
End If
 
If pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity <> 0 Then
oLoad.EntityKey = pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity
'send it without transaction if first record, and send in with transaction for subsequent records. Otherwise, it will deadlock
'Need to fill in the Load Object so that it is populated when sending Cancel Preforecast to TFCS
If icount = 1 Then
     Call oLoad.GetLoadByEntityKey(CStr(pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity))
Else
     Call oLoad.GetLoadByEntityKey(CStr(pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity), , lTransaction)
End If
End If
'Wrap the data
oLoad.OutboundFRDTrailerGenNR = pInst.PDInfo.hPDHll.Item(icount).lCurForTlrEntity
'oLoad.OriginCountryCode = pInst.PDInfo.hPDHll.Item(icount).szOriginCn
oLoad.OriginName = pInst.PDInfo.hPDHll.Item(icount).szOrigin
oLoad.OriginCountryCode = pInst.PDInfo.hPDHll.Item(icount).szOriginCn
oLoad.DestinationName = pInst.PDInfo.hPDHll.Item(icount).szDestin
oLoad.DestinationSortTypeCode = pInst.PDInfo.hPDHll.Item(icount).szDestinSrt
oLoad.SequenceNumber = Format$(CStr(pInst.PDInfo.hPDHll(icount).iSequenceNr), "00")
oLoad.DestinationCountryCode = pInst.PDInfo.hPDHll.Item(icount).szDestinCn
If pInst.PDInfo.hPDHll.Item(icount).dtLdCrtDt <> 0 Then
    oLoad.CreateDate = FmtDbDate(pInst.PDInfo.hPDHll.Item(icount).dtLdCrtDt)
End If
oLoad.EqpDirectionTypeCode = "O"
oLoad.SchTlrEqpGenNr = pInst.PDInfo.hPDHll.Item(icount).lSchTlrEqpGenNr
'HFCS_INDICATOR_TRUE
oLoad.PredispatchJob = txtJobNr.Text
If Len(pInst.PDInfo.hPDHll.Item(icount).szCurPDJobDomCn) > 0 Then
   oLoad.OutboundActSLOCCountryCD = pInst.PDInfo.hPDHll.Item(icount).szCurPDJobDomCn
Else
   oLoad.OutboundActSLOCCountryCD = pInst.PDInfo.szPDJobDomCn
End If
If Len(pInst.PDInfo.hPDHll.Item(icount).szCurPDJobDom) > 0 Then
    oLoad.OutboundActSLOC = pInst.PDInfo.hPDHll.Item(icount).szCurPDJobDom
Else
    oLoad.OutboundActSLOC = pInst.PDInfo.szPDJobDomSlic
End If
oLoad.PredispatchDate = Date
oLoad.PredispatchTime = Format(Now, "HH:MM:SS")
oLoad.OriginSortTypeCode = pInst.PDInfo.hPDHll.Item(icount).szOriginSrt
 
'KG - WR00590, if ELOC in structure is not set, use ELOC Country variable
'from PDInfo
'If Len(pInst.PDInfo.hPDHll.Item(icount).szElocCn) > 0 Then
'    oLoad.PredispatchELOCCnyCd = pInst.PDInfo.hPDHll.Item(icount).szElocCn
'Else
oLoad.PredispatchELOCCnyCd = pInst.PDInfo.szPDElocCn
'End If
'*** end WR00590
      
oLoad.PredispatchedSegmentSystemOID = pInst.PDInfo.SegmentSystemNumberOID
oLoad.JobNumberSystemOID = pInst.PDInfo.JobSystemNumberOID
oLoad.PredispatchELOC = Left$(txtEloc.Text, 5)
oLoad.LoadMsgSeqNumber = pInst.PDInfo.hPDHll.Item(icount).lLdMsgId
oLoad.TrailerMsgSeqNumber = pInst.PDInfo.hPDHll.Item(icount).lTlrMsgId
oLoad.TrailerNumber = pInst.PDInfo.hPDHll.Item(icount).szTlrNr
oLoad.TrailerTypeCode = pInst.PDInfo.hPDHll.Item(icount).szEqpTyp
If pInst.PDInfo.hPDHll.Item(icount).szMultLdIr Then
   oLoad.HasMultipleLoads = "1"
Else
   oLoad.HasMultipleLoads = "0"
End If
oLoad.PredispatchDriver = txtDriverName.Text
oLoad.PredispatchTractor = txtTractorNumber.Text
 
oLoad.OutboundVehicleEntityKey = pInst.PDInfo.lPDFopVehEntKey  'pInst.PDInfo.hPDHll.Item(icount).lSchdVehEntity
oLoad.DollyKey = pInst.PDInfo.hPDHll.Item(icount).lCurFopDolEntity
oLoad.BayNumber = ctlBayNr(icount - 1).Text
oLoad.LoadCode = pInst.PDInfo.hPDHll.Item(icount).szLdCd
oLoad.Position = pInst.PDInfo.hPDHll.Item(icount).szTlrPosCd
oLoad.RouteCode = pInst.PDInfo.hPDHll.Item(icount).szRtId
oLoad.DriverNotified = pInst.PDInfo.szFdrDvrInfNtfIr
oLoad.PackageCount = pInst.PDInfo.hPDHll.Item(icount).lPieces
oLoad.PackagePercentage = pInst.PDInfo.hPDHll.Item(icount).iPercent
oLoad.DollyName = txtDolly(icount - 1).Text
If chkExtraDolly(0).Enabled = True And icount = 1 Then  'icount = 1 Then
    oLoad.ExtraDolly = CStr(chkExtraDolly(0).Value)
ElseIf chkExtraDolly(1).Enabled = True And icount = 2 Then 'icount = 2 Then
    oLoad.ExtraDolly = CStr(chkExtraDolly(1).Value)
Else
    oLoad.ExtraDolly = "0"
End If
 
If pInst.PDInfo.hPDHll.Item(icount).lSchFdrEntKey <> -1 Then
    oLoad.PreDispatchScheduledEntityKey = pInst.PDInfo.hPDHll.Item(icount).lSchFdrEntKey
End If
 
oLoad.OutboundFeederScheduleWeekEndingDate = ctlWkendingDate.Value
oLoad.OutboundFeederScheduleDOW = cboDow.ListIndex
'oLoad.OutboundSchedDate = Format(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy")
 
oLoad.DepartureDate = Format(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy")
oLoad.DepartureTime = Format(pInst.PDInfo.FeederScheduleEndTime, "hh:mm")
 
'send predispatch preforecast
oLoad.SendPredispatchPreforecast = pInst.PDInfo.hPDHll(icount).SendPredispatchPreforecast
 
If Not bFromGtF8 Then
   bCreateOutboundRecord(icount - 1) = oClsDb.CheckOutboundRecord(icount, lTransaction)
   'TT00860
   If bCreateOutboundRecord(icount - 1) Or (IsNull(pInst.PDInfo.hPDHll.Item(icount).AssignedPDOID) _
      And bPredispatchEmpty(icount - 1) = True And pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity > 0) Then
      sTableNa = "TFORTLR"
      oLoad.OutboundFRDTrailerGenNR = oLoad.GlobalGetEntityKey(m_HFCSDB, Me, sTableNa)
     pInst.PDInfo.hPDHll.Item(icount).lCurForTlrEntity = oLoad.OutboundFRDTrailerGenNR
      'TT00860
      If bPredispatchEmpty(icount - 1) = True Then
         pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity = 0
         oLoad.EntityKey = 0
      End If
      bResult = oLoad.InsertLoadInfoForPreDispatch(txtDolly.Item(icount - 1).Text, , lTransaction)
      bMatchingLoadFound(icount - 1) = False
   End If
  
   If lOrgFRDKey(icount - 1) = oLoad.OutboundFRDTrailerGenNR And oLoad.EntityKey <> 0 Then
      bSameFrdKey = True
   End If
 
   bResult = oLoad.UpdateLoadInfoForPreDispatch(txtDolly.Item(icount - 1).Text, , lTransaction, bSameFrdKey)
     
End If
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD AddPredispatchDriverData - End")
End If
Exit Function
 
ErrorHandler:
glErrNum = 400
gsError = "AddPredispatchDriverData"
Call oProc.update_error_object(Me, gsError)
oProc.update_error_object Nothing, sFnName
With Err
    'Handle the error here as you normally would
    oErrorObject.error_routine oEventlog.GlobalEvents, _
                                .Number, _
                                .Description, _
                                oProc, _
                                ERROR_MSG, _
                                FEEDER_DISPATCH_DRIVER, _
                                NO_POP_UP
    .Raise LOAD_ERRORS.LOADS_OBJ_UNKNOWN_ERROR, _
            .Source, _
            .Description, _
            .HelpFile, _
            .HelpContext
End With
'Unload Me
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
End Function
 
Public Function AddPredispatchDriverDataForArrivals(lTracNr As Long, icount As Integer, _
                                            Optional UpdDriverNotifyOnly As Boolean = False, _
                                            Optional iArrIndex As Integer = 0) As Boolean
On Error GoTo ErrorHandler
Dim bResult As Boolean
Dim sFnName As String
Dim lTransNr As Long
Dim sSQL As String
Dim lCurForTlrEntity As Long
Dim sSequenceNr As String
Dim bInsert As Boolean
Dim i As Integer
Dim sTableNa As String
Dim j As Integer
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD AddPredispatchDriverDataForArrivals - Begin")
End If
 
bResult = True
 
sFnName = "AddPredispatchDriverDataForArrivals"
 
If iArrIndex = 0 Then
    i = iBayIndex + 1
Else
    'find the correct load in the arrival structure, order can be different on
    'predispatch screen
    For j = 1 To iPredispatch
       If icount = raArrivalData(j).iSelected + 1 Then
           Exit For
       End If
    Next j
    If j = iArrIndex Then
      i = iArrIndex
    Else
      i = j
    End If
End If
 
'Wrap data for Pre-dispatch
If pInst.PDInfo.hPDHll.Item(icount).lCurForTlrEntity <> 0 Then
    raArrivalData(i).lForTlrEntKey = pInst.PDInfo.hPDHll.Item(icount).lCurForTlrEntity
    oLoad.EntityKey = raArrivalData(i).lFopTlrEntKey
    'send it without transaction if first record, and send in with transaction for subsequent records. Otherwise, it will deadlock
    'Need to fill in the Load Object so that it is populated when sending Cancel Preforecast to TFCS
    If i = 1 Then
        Call oLoad.GetLoadByEntityKey(CStr(raArrivalData(i).lFopTlrEntKey))
    Else
        Call oLoad.GetLoadByEntityKey(CStr(raArrivalData(i).lFopTlrEntKey), , lTracNr)
    End If
End If
If bPredispatchEmptyWithTrailer(icount - 1) = True Then
    oLoad.PredispatchSequenceEdited = True
End If
oLoad.OutboundFRDTrailerGenNR = raArrivalData(i).lForTlrEntKey
oLoad.OriginCountryCode = raArrivalData(i).sLdOrgCnyCd
oLoad.OriginName = raArrivalData(i).sLdOrgSlcAbrNa
oLoad.DestinationCountryCode = raArrivalData(i).sLdDtnCnyCd
oLoad.DestinationName = raArrivalData(i).sLdDtnSlcAbrNa
oLoad.DestinationSortTypeCode = raArrivalData(i).sLdDtnSrtTypCd
If Len(CStr(raArrivalData(i).sLdSeqNr)) = 0 Then
   oLoad.SequenceNumber = "00"
Else
   oLoad.SequenceNumber = CStr(raArrivalData(i).sLdSeqNr)
End If
     
oLoad.CreateDate = raArrivalData(i).dtLdCrtDt
oLoad.EqpDirectionTypeCode = "O"
oLoad.OutboundSchedDate = raArrivalData(i).dtDueDt
oLoad.PredispatchJob = txtJobNr.Text
oLoad.PredispatchDate = Date
oLoad.PredispatchTime = Format(Now, "HH:MM:SS")
oLoad.OriginSortTypeCode = raArrivalData(i).sActOrgSrtTypCd
oLoad.OutboundActSLOCCountryCD = raArrivalData(1).sFdrJobDmcCnyCd
oLoad.OutboundActSLOC = raArrivalData(1).sFdrJobDmcSAB
If Len(txtEloc.Text) = 8 Then
   oLoad.PredispatchELOCCnyCd = Right$(txtEloc.Text, 2)
Else
   oLoad.PredispatchELOCCnyCd = raArrivalData(1).sFdrJobDmcCnyCd
End If
      
oLoad.PredispatchELOC = Left$(txtEloc.Text, 5)
oLoad.PredispatchedSegmentSystemOID = pInst.PDInfo.SegmentSystemNumberOID
oLoad.JobNumberSystemOID = pInst.PDInfo.JobSystemNumberOID
oLoad.LoadCode = raArrivalData(i).sActLdCd
oLoad.LoadMsgSeqNumber = raArrivalData(i).lActLdMsgId
oLoad.TrailerMsgSeqNumber = raArrivalData(i).lActTlrMsgId
oLoad.TrailerNumber = pInst.PDInfo.hPDHll.Item(icount).szTlrNr
oLoad.HasMultipleLoads = raArrivalData(i).bActMulLdIr
oLoad.TrailerMsgSeqNumber = raArrivalData(i).lActTlrMsgId
oLoad.PredispatchDriver = txtDriverName.Text
oLoad.PredispatchTractor = txtTractorNumber.Text
oLoad.EntityKey = raArrivalData(i).lFopTlrEntKey
oLoad.OutboundVehicleEntityKey = raArrivalData(i).lFopVehEntKey
oLoad.PredispatchedSegmentSystemOID = raArrivalData(i).sSegSysNrOID
        
oLoad.DollyKey = pInst.PDInfo.hPDHll.Item(icount).lCurFopDolEntity
oLoad.DollyName = txtDolly(icount - 1).Text
oLoad.BayNumber = raArrivalData(i).sBayNr
oLoad.RouteCode = raArrivalData(i).sActRteTypCd
oLoad.DriverNotified = pInst.PDInfo.szFdrDvrInfNtfIr
     
oLoad.OutboundFeederScheduleWeekEndingDate = ctlWkendingDate.Validate
oLoad.OutboundFeederScheduleDOW = cboDow.ListIndex
 
oLoad.DepartureDate = Format(pInst.PDInfo.FeederScheduleEndTime, "mm/dd/yyyy")
oLoad.DepartureTime = Format(pInst.PDInfo.FeederScheduleEndTime, "hh:mm")
 
oLoad.SendPredispatchPreforecast = pInst.PDInfo.hPDHll.Item(icount).SendPredispatchPreforecast
If chkExtraDolly(0).Enabled = True And icount = 1 Then  'icount = 1 Then
    oLoad.ExtraDolly = CStr(chkExtraDolly(0).Value)
ElseIf chkExtraDolly(1).Enabled = True And icount = 2 Then 'icount = 2 Then
    oLoad.ExtraDolly = CStr(chkExtraDolly(1).Value)
Else
    oLoad.ExtraDolly = "0"
End If
     
If Not bFromGtF8 Then
   bCreateOutboundRecord(icount - 1) = oClsDb.CheckOutboundRecord(icount, lTracNr)
   'TT00860
   If bCreateOutboundRecord(icount - 1) Or (IsNull(pInst.PDInfo.hPDHll.Item(icount).AssignedPDOID) _
      And bPredispatchEmpty(icount - 1) = True And pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity > 0) Then
      sTableNa = "TFORTLR"
      oLoad.OutboundFRDTrailerGenNR = oLoad.GlobalGetEntityKey(m_HFCSDB, Me, sTableNa)
      'TT00860
      If bPredispatchEmpty(icount - 1) = True Then
         pInst.PDInfo.hPDHll.Item(icount).lCurFopTlrEntity = 0
         oLoad.EntityKey = 0
      End If
      bResult = oLoad.InsertLoadInfoForPreDispatch(txtDolly.Item(icount - 1).Text, , lTracNr)
   End If
   bResult = oLoad.UpdateLoadInfoForPreDispatch(txtDolly.Item(icount - 1).Text, , lTracNr, True)
End If
 
AddPredispatchDriverDataForArrivals = bResult
'Set oLoad = Nothing
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD AddPredispatchDriverDataForArrivals - End")
End If
Exit Function
 
ErrorHandler:
glErrNum = 400
gsError = "AddPredispatchDriverDataForArrivals"
Call oProc.update_error_object(Me, gsError)
'oProc.update_error_object Nothing, sFnName
AddPredispatchDriverDataForArrivals = False
With Err
    'Handle the error here as you normally would
    oErrorObject.error_routine oEventlog.GlobalEvents, _
                                .Number, _
                                .Description, _
                                oProc, _
                                ERROR_MSG, _
                                FEEDER_DISPATCH_DRIVER, _
                                NO_POP_UP
    .Raise LOAD_ERRORS.LOADS_OBJ_UNKNOWN_ERROR, _
            .Source, _
            .Description, _
            .HelpFile, _
            .HelpContext
End With
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'Unload Me
End Function
 
Public Function UpdatePdDataNew() As Integer
 
On Error GoTo ErrorExit
 
' Save data to the database
Dim i As Integer
Dim iPdCount As Integer
Dim bResult As Boolean
Dim oTempDolly As Dolly
Dim iRc As Integer
Dim objCollOfRa As Collection
Dim objTmpStruct As PdStruct
Dim lTransNr As Long
Dim vTmpPdRecData As Variant
Dim j As Integer
Dim sTableNa As String
Dim iArrIndex As Integer
Dim lFrdKey(3) As Long
Dim lFopKey(3) As Long
Dim lDollyKey(3) As Long
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD UpdatePdDataNew - Begin")
End If
 
Set m_HFCSDB = oClsDb.DBClass
Set objCollOfRa = New Collection
 
Dim bReturnCode As Boolean
   
iArrIndex = 0
For i = 0 To 2
    lFrdKey(i) = 0
    lFopKey(i) = 0
    lDollyKey(i) = 0
Next i
lTransNr = m_HFCSDB.Start_Transaction
For i = 1 To iPredispatch
    If bPredispatchEmpty(i - 1) Then
        Call UpdatePredispatchEmptyData(i - 1)
    End If
    'TT0749
    If bPredispatchEmptyWithTrailer(i - 1) = True Or bPredispatchEmpty(i - 1) = True Then
        If grdBayDetails(i - 1).Col = GRD1_COL_TYPE And bPredispatchEmpty(i - 1) = True Then
            grdBayDetails(i - 1).Split = GRD1_SPL_LOAD
            grdBayDetails(i - 1).Col = GRD1_COL_DEST
        End If
       
        Call grdBayDetails_BeforeRowColChange(i - 1, False)
       
        If (grdBayDetails(i - 1).Col = GRD1_COL_DEST) Then
            grdBayDetails(i - 1).Split = GRD1_SPL_LOAD
            grdBayDetails(i - 1).Col = GRD1_COL_DS
            Call grdBayDetails_BeforeRowColChange(i - 1, False)
        End If
       
        If ctlWarningPopup.WarningVisible = True Then
            If grdBayDetails(i - 1).Enabled = False Then grdBayDetails(i - 1).Enabled = True
            grdBayDetails(i - 1).SetFocus
            UpdatePdDataNew = DB_GRID_ERROR
            m_HFCSDB.RollBack_Transaction lTransNr
            m_HFCSDB.End_Transaction lTransNr
            If g_iDebug = 13 Then
               Call InfoLog("frmPDD UpdatePdDataNew - End - If ctlWarningPopup.WarningVisible = True ")
            End If
            Exit Function
        ElseIf Trim$(grdBayDetails(i - 1).Columns(GRD1_COL_RI).Value) = "" Then
            If grdBayDetails(i - 1).Enabled = False Then grdBayDetails(i - 1).Enabled = True
            grdBayDetails(i - 1).SetFocus
            If grdBayDetails(i - 1).Col <> GRD1_COL_RI Then
                grdBayDetails(i - 1).Split = 3
                grdBayDetails(i - 1).Col = GRD1_COL_RI
            End If
            Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(i - 1).hwnd, "Route Code not entered, enter route code", 2250, _
                                                 grdBayDetails(i - 1).RowTop(grdBayDetails(i - 1).Row), _
                                                 grdBayDetails(i - 1).Columns(GRD1_COL_RI).Left, _
                                                 grdBayDetails(i - 1).Columns(GRD1_COL_RI).Width, _
                                                 grdBayDetails(i - 1).RowHeight)
 
            UpdatePdDataNew = DB_GRID_ERROR
            m_HFCSDB.RollBack_Transaction lTransNr
            m_HFCSDB.End_Transaction lTransNr
            If g_iDebug = 13 Then
               Call InfoLog("frmPDD UpdatePdDataNew - End - ElseIf Trim$(grdBayDetails(i - 1).Columns(GRD1_COL_RI).Value) = """)
            End If
            Exit Function
        ElseIf grdBayDetails(i - 1).Col = GRD1_COL_DS Then
            grdBayDetails(i - 1).Split = 3
            grdBayDetails(i - 1).Col = GRD1_COL_RI
            Call grdBayDetails_BeforeRowColChange(i - 1, False)
            If (bBayCompeleted(i - 1) = False And bBalNrValidated(i - 1) = False) Or ctlWarningPopup.WarningVisible = True Then
                'grdBayDetails(i - 1).SetFocus
                grdBayDetails(i - 1).Split = 0
                grdBayDetails(i - 1).Col = 0
                If ctlBayNr(i - 1).Enabled = False Then ctlBayNr(i - 1).Enabled = True
                ctlBayNr(i - 1).SetFocus
               ' If Len(ctlBayNr(i - 1).Text) > 0 Then DoEvents
                UpdatePdDataNew = DB_GRID_ERROR
                m_HFCSDB.RollBack_Transaction lTransNr
                m_HFCSDB.End_Transaction lTransNr
                If g_iDebug = 13 Then
                   Call InfoLog("frmPDD UpdatePdDataNew - End - ElseIf grdBayDetails(i - 1).Col = GRD1_COL_DS ")
                End If
                Exit Function
            Else
                Call pInst.UpdatePredispatchWithEmpty(i - 1)
                grdBayDetails(i - 1).Columns(GRD1_COL_OS).Value = pInst.PDInfo.hPDHll(i).szOriginSrt
                grdBayDetails(i - 1).Columns(GRD1_COL_SEQ).Value = Format$(pInst.PDInfo.hPDHll(i).iSequenceNr, "00")
            End If
        Else
           Call pInst.UpdatePredispatchWithEmpty(i - 1)
            grdBayDetails(i - 1).Columns(GRD1_COL_OS).Value = pInst.PDInfo.hPDHll(i).szOriginSrt
'            pInst.PDInfo.hPDHll(i).iSequenceNr = CInt(grdBayDetails(i - 1).Columns(GRD1_COL_SEQ).Value)
            grdBayDetails(i - 1).Columns(GRD1_COL_SEQ).Value = Format$(pInst.PDInfo.hPDHll(i).iSequenceNr, "00")
        End If
    End If
    'end TT0749
    Set oLoad = New HFCSLoadObject.HFCSLOAD
    oLoad.SpecifyConnection m_HFCSDB
   
 
    If pInst.bArriveFlag And Not bArrivalTractorOnly And bFromGTForm(i - 1) Then
        iArrIndex = iArrIndex + 1
        'Create the outbound record only for arrivals
        bReturnCode = AddPredispatchDriverDataForArrivals(lTransNr, i, , iArrIndex)
 
    Else  'For Pre-dispatch itself
        'Create the outbound record only for non arrivals
        bReturnCode = AddPredispatchDriverData(lTransNr, i)
    End If
    
    If bMatchingLoadFound(iPredispatch - 1) Then
       bReturnCode = UpdateTFORTLR(lTransNr, i)
    End If
   
   'For arrival with F8
    If oClsDb.CheckOutboundRecord(i, lTransNr) Then
       sTableNa = "TFORTLR"
       oLoad.OutboundFRDTrailerGenNR = oLoad.GlobalGetEntityKey(m_HFCSDB, Me, sTableNa)
       sPDPinfo = sPDPinfo & "INSERT|"
    ElseIf Not bCreateOutboundRecord(i - 1) And bFromGtF8 Then
       sPDPinfo = sPDPinfo & "UPDATE|"
    ElseIf bFromGtF8 Then
       sPDPinfo = sPDPinfo & "UPDATE|"
       oLoad.OutboundFRDTrailerGenNR = pInst.PDInfo.hPDHll.Item(i).lCurForTlrEntity
       bCreateOutboundRecord(i - 1) = False
    End If
    lFrdKey(i - 1) = oLoad.OutboundFRDTrailerGenNR
    lFopKey(i - 1) = oLoad.EntityKey
    lDollyKey(i - 1) = CLng(oLoad.OutboundVehicleEntityKey)
   
    If bFromGtF8 Then
       Set oLoadInfo(i - 1) = oLoad
    Else
       If pInst.PDInfo.hPDHll.Item(i).szMultLdIr = "1" Then
           oLoad.HasMultipleLoads = True
       End If
      
       Set oColLoads(i - 1) = oLoad
'      If Not oForecast Is Nothing Then
'       Set oColForecast(i - 1) = oForecast
'      End If
      
    End If
   
    Set oLoad = Nothing
Next i
 
If Not bFromGtF8 Then
    Set oLoad = New HFCSLoadObject.HFCSLOAD
    oLoad.PredispatchedSegmentSystemOID = pInst.PDInfo.SegmentSystemNumberOID
    oLoad.SpecifyConnection oClsDb.DBClass
'    Call oLoad.SetPDFopKey(lFopKey, lDollyKey)
    bReturnCode = oLoad.RemoveOldPredispatch(lFrdKey, lTransNr)
    Set oLoad = Nothing
End If
Set objTmpStruct = Nothing
   
m_HFCSDB.Commit_Transaction lTransNr
m_HFCSDB.End_Transaction lTransNr
   
Set m_HFCSDB = Nothing
If g_iDebug = 13 Then
   Call InfoLog("frmPDD UpdatePdDataNew - End")
End If
Exit Function
 
ErrorExit:
glErrNum = 400
gsError = "UpdatePdDataNew"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbNormal
 
m_HFCSDB.RollBack_Transaction lTransNr
m_HFCSDB.End_Transaction lTransNr
 
If oErrorObject.error_routine(oEventlog.FeederShell, _
                             IIf(Err.Number <> 0, Err.Number, glErrNum), _
                             Err.Description & " Module:" & gsError, _
                             oProc, _
                             ERROR_MSG, _
                             FEEDER_DISPATCH_DRIVER) Then
   
    Set oEventlog = Nothing
    MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
  '  Unload Me
 
End If
 
End Function
Public Function UpdateTFORTLR(lTransaction As Long, icount As Integer, _
                              Optional UpdDriverNotifyOnly As Boolean = False) As Boolean
 
Dim bResultRecord As Boolean
Dim sFnName As String
Dim lTransNr As Long
Dim sSQL As String
Dim lCurForTlrEntity As Long
Dim sSequenceNr As String
Dim bInsert As Boolean
Dim sTable As String
Dim rsCheckOrgFrdKey As ADODB.Recordset
Dim bUpdatePredispatch As Boolean
 
On Error GoTo ErrorHandler
If g_iDebug = 13 Then
   Call InfoLog("frmPDD UpdateTFORTLR - Begin")
End If
 
If (m_HFCSDB Is Nothing) Then
    Set m_HFCSDB = oClsDb.DBClass
End If
 
sFnName = "UpdateTFORTLR"
bResultRecord = False
bUpdatePredispatch = True
 
If lOrgFRDKey(icount - 1) <> pInst.PDInfo.hPDHll.Item(icount).lCurForTlrEntity And lOrgFRDKey(icount - 1) > 0 Then
    Set rsCheckOrgFrdKey = oClsDb.CheckOrgFrdKeyAssignedToAnotherJob(lOrgFRDKey(icount - 1), lTransaction)
   
    'TT03284
    If Not (rsCheckOrgFrdKey Is Nothing) Then
        If (rsCheckOrgFrdKey.BOF And rsCheckOrgFrdKey.EOF) Then
            bUpdatePredispatch = oClsDb.CheckOrgFrdKeyHasTrailer(lOrgFRDKey(icount - 1), lTransaction)
        Else
            If Not IsNull(rsCheckOrgFrdKey.Fields("ASN_FDR_JOB_NR").Value) Then
                If (rsCheckOrgFrdKey.Fields("FRD_TLR_EQP_GEN_NR").Value = lOrgFRDKey(icount - 1) _
                    And Trim$(rsCheckOrgFrdKey.Fields("ASN_FDR_JOB_NR").Value) = Trim$(txtJobNr.Text)) Then
                    bUpdatePredispatch = oClsDb.CheckOrgFrdKeyHasTrailer(lOrgFRDKey(icount - 1), lTransaction)
                Else
                    bUpdatePredispatch = False
                End If
            Else
                bUpdatePredispatch = oClsDb.CheckOrgFrdKeyHasTrailer(lOrgFRDKey(icount - 1), lTransaction)
            End If
        End If
    Else
        bUpdatePredispatch = oClsDb.CheckOrgFrdKeyHasTrailer(lOrgFRDKey(icount - 1), lTransaction)
    End If
   
    'TT03284
    If bUpdatePredispatch = True Then
        'Update the FOP Key of the one which is on the Bay to Null.
        bResultRecord = False
        UpdateTFORTLR = oLoad.UpdateForPreDispatch(lOrgFRDKey(icount - 1), txtDolly.Item(icount - 1).Text, _
                    bResultRecord, lTransaction)
    End If
 
 
 
'when this occurs, there is an outbound record, but the load name for the new schedule does
'not exist. Set the old outbound record to NULL, and insert the new record.
ElseIf lOrgFRDKey(icount - 1) = 0 And bCreateOutboundRecord(icount - 1) = False Then
       sTable = "TFORTLR"
    lOrgFRDKey(icount - 1) = oLoad.GlobalGetEntityKey(m_HFCSDB, Me, sTable)
    bResultRecord = True
    UpdateTFORTLR = oLoad.UpdateForPreDispatch(lOrgFRDKey(icount - 1), txtDolly.Item(icount - 1).Text, _
                bResultRecord, lTransaction)
End If
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD UpdateTFORTLR - End")
End If
 
Exit Function
 
ErrorHandler:
 
glErrNum = 400
gsError = "UpdateTFORTLRl"
Call oProc.update_error_object(Me, gsError)
oProc.update_error_object Nothing, sFnName
With Err
   'Handle the error here as you normally would
   oErrorObject.error_routine oEventlog.GlobalEvents, _
               .Number, _
               .Description, _
                oProc, _
                ERROR_MSG, _
                FEEDER_DISPATCH_DRIVER, _
                NO_POP_UP
   .Raise LOAD_ERRORS.LOADS_OBJ_UNKNOWN_ERROR, _
       .Source, _
       .Description, _
       .HelpFile, _
       .HelpContext
End With
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'Unload Me
End Function
'<WR01024><ADID: dmx7plm> - Start
Private Function ClosefrmPDD() As Boolean
On Error GoTo Error_Handler
 
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ClosefrmPDD - Begin")
    End If
 
    Dim lHwnd As Long
    ClosefrmPDD = False
   
    If bClickedF4 = True Then lHwnd = GetWindowHandle(OGlobalFormNames.StartBaySearch)
    If bClickedF5 = True Then lHwnd = GetWindowHandle(OGlobalFormNames.QuickSearchF5)
   
    If (lHwnd = 0) Then
        Unload Me
        ClosefrmPDD = True
    Else
        MsgBox LoadResString(1107), vbOKOnly + vbInformation, App.Title
    End If
   
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ClosefrmPDD - End")
    End If
    Exit Function
   
Error_Handler:
glErrNum = 400
gsError = "ClosefrmPDD"
Call oProc.update_error_object(Me, gsError)
 
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
      '  Unload Me
End If
End Function
'<WR01024><ADID: dmx7plm> - End
 
 
Private Sub UpdatePredispatchEmptyData(Index As Integer)
Dim sSab As String
Dim sCny As String
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD UpdatePredispatchEmptyData - Begin - Index: " & Index)
End If
 
pInst.PDInfo.hPDHll(Index + 1).szEqpTyp = grdBayDetails(Index).Columns(GRD1_COL_TYPE).Value
pInst.PDInfo.hPDHll(Index + 1).szRtId = grdBayDetails(Index).Columns(GRD1_COL_RI).Value
 
SplitSabCny grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, sSab, sCny, pInst.PDInfo.szCurrentCny
pInst.PDInfo.hPDHll(Index + 1).szDestin = sSab
pInst.PDInfo.hPDHll(Index + 1).szDestinCn = sCny
pInst.PDInfo.hPDHll(Index + 1).szDestinSrt = grdBayDetails(Index).Columns(GRD1_COL_DS).Value
 
If g_iDebug = 13 Then
   Call InfoLog("frmPDD UpdatePredispatchEmptyData - End - Index: " & Index)
End If
End Sub
 
 
Public Function PerformFieldLevelValidation(Index As Integer) As Boolean
  Dim sGroupCode            As String
  Dim iCounter              As Integer
  Dim bFound                As Boolean
  Dim sOriginSite           As String
  Dim sDestSite             As String
  Dim sSlocName             As String
  Dim sElocName             As String
  ' SCR#3504
  Dim sSlic                 As String
  Dim sCountryCode          As String
  Dim sLocalCn              As String
  Dim sOpType               As String
  Dim sOrigSlic             As String
  Dim bValidRouting         As Boolean
  Dim iPosition As Integer
  Dim sTlrTyp               As String
  Dim rsTrailerTypes        As ADODB.Recordset
  Dim sOriginalRouteCd      As String
  Dim iRc                   As Integer
  Dim bValidateLoadRouting   As Boolean
  Dim bInvalidLoadRouting       As Boolean
  Dim sElocCnyCd            As String
  Dim oLoadRoutings As HFCSLoadRoutingCodes.LoadRoutingCodes
  On Error GoTo Error_Handler
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD PerformFieldLevelValidation - Begin: Index = " & Index)
    End If
  'Assume failure
  PerformFieldLevelValidation = False
  bValidateLoadRouting = False
    Set oLoadRoutings = New HFCSLoadRoutingCodes.LoadRoutingCodes
 
  If ctlWarningPopup.WarningVisible = True Then
    ctlWarningPopup.ClearWarning
  End If
  'grdBayDetails(Index).EditActive = False
  sLocalCn = pInst.PDInfo.szCurrentCny
 
  Select Case grdBayDetails(Index).Col
    Case GRD1_COL_DEST
        'Lets make sure they at least key in something.  We really cannot validate until we now
        'what they are putting in for the origin sort.
        If Trim$(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Value) <> vbNullString Then
           iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, "/")
          
            If iPosition = 0 Then
                iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, " ")
            End If
            
            If iPosition > 0 And Len(Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value)) >= iPosition Then
                sCountryCode = Mid$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, iPosition + 1, 2)
                sDestSite = Mid$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, 1, 5)
            Else
                sDestSite = Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value)
                sCountryCode = pInst.PDInfo.szCurrentCny
            End If
             
            bFound = oClsDb.GetSlicData(sCountryCode, sDestSite, sSlic, , , , sOpType)
            
            If bPredispatchEmpty(Index) Then
                If Index > 0 Then
                    If Index = 1 Then
                         If pInst.PDInfo.hPDHll(Index + 1).SameLoadName = True And grdBayDetails(Index).Columns(GRD1_COL_DEST).DataChanged = True Then
                              If sDestSite <> Mid$(Trim$(grdBayDetails(0).Columns(GRD1_COL_DEST).Value), 1, 5) Then
                                  pInst.PDInfo.hPDHll(Index + 1).SameLoadName = False
                              End If
                         End If
                    End If
                   
                    If Index = 2 Then
                         If pInst.PDInfo.hPDHll(Index + 1).SameLoadName = True And grdBayDetails(Index).Columns(GRD1_COL_DEST).DataChanged = True Then
                              If sDestSite <> Mid$(Trim$(grdBayDetails(0).Columns(GRD1_COL_DEST).Value), 1, 5) And _
                                 sDestSite <> Mid$(Trim$(grdBayDetails(0).Columns(GRD1_COL_DEST).Value), 1, 5) Then
                                  pInst.PDInfo.hPDHll(Index + 1).SameLoadName = False
                              End If
                         End If
                    End If
                End If
            End If
           
            If bFound Then
                pInst.PDInfo.hPDHll(Index + 1).szDestinCn = sCountryCode
                pInst.PDInfo.hPDHll(Index + 1).szDestin = sDestSite
            End If
           
            If Trim$(sDestSite) = "EMPTY" Or ValidateSLIC(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, sLocalCn, sSlic, sCountryCode) Then
              If sCountryCode = sLocalCn Then
                grdBayDetails(Index).Columns(GRD1_COL_DEST).Value = sSlic
              End If
             
              If Trim$(sOpType) = "CPU" Or Trim$(sOpType) = "CPX" Then
                  grdBayDetails(Index).Columns(GRD1_COL_DS).Value = "C"
              Else
                  grdBayDetails(Index).Columns(GRD1_COL_DS).Value = "E"
              End If
            Else
              If grdBayDetails(Index).Visible Then
                grdBayDetails(Index).Enabled = True
                grdBayDetails(Index).SetFocus
              End If
             
              grdBayDetails(Index).Columns(GRD1_COL_DS).Value = "C"
            End If
        Else
         
           'Make sure focus is on the Bay grid.
            If grdBayDetails(Index).Visible Then
                grdBayDetails(Index).Enabled = True
                grdBayDetails(Index).SetFocus
            End If
         
            Call ctlWarningPopup.ShowWarning(grdBayDetails(Index).hwnd, LoadResString(sInvalidDestinationSLIC), 2250, _
                                                    grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                                    grdBayDetails(Index).Columns(GRD1_COL_DEST).Left, _
                                                    grdBayDetails(Index).Columns(GRD1_COL_DEST).Width, _
                                                    grdBayDetails(Index).RowHeight)
            Set oLoadRoutings = Nothing
            If g_iDebug = 13 Then
               Call InfoLog("frmPDD PerformFieldLevelValidation - End - If bPredispatchEmpty(Index) - Index = " & Index)
            End If
            Exit Function
        End If
        If grdBayDetails(Index).Columns(GRD1_COL_DEST).DataChanged = True Then
            grdBayDetails(Index).Columns(GRD1_COL_RI).Value = "G"
        End If
 
    Case GRD1_COL_DS
          iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, "/")
          
          If iPosition = 0 Then
              iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, " ")
          End If
          
          If iPosition > 0 And Len(Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value)) >= iPosition Then
              sCountryCode = Mid$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, iPosition + 1, 2)
              sDestSite = Mid$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, 1, 5)
          Else
              sDestSite = Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value)
              sCountryCode = pInst.PDInfo.szCurrentCny
          End If
         
          bFound = oClsDb.GetSlicData(sCountryCode, sDestSite, "", , , , sOpType)
         
          If grdBayDetails(Index).Columns(GRD1_COL_DS).DataChanged = True _
             And Trim$(grdBayDetails(Index).Columns(GRD1_COL_DS).Value) <> "C" _
             And Trim$(grdBayDetails(Index).Columns(GRD1_COL_DS).Value) <> "E" Then
                          
                 
             If grdBayDetails(Index).Visible Then
                grdBayDetails(Index).Enabled = True
                grdBayDetails(Index).SetFocus
             End If
            
             Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, LoadResString(sInvalidDestinationSort), 2250, _
                                              grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                              grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Left, _
                                              grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Width, _
                                              grdBayDetails(Index).RowHeight)
             Set oLoadRoutings = Nothing
             If g_iDebug = 13 Then
                 Call InfoLog("frmPDD PerformFieldLevelValidation - End - If grdBayDetails(Index).Columns(GRD1_COL_DS).DataChanged = True: GRD1_COL_DS = " & GRD1_COL_DS & " - Index: " & Index)
             End If
             Exit Function
          ElseIf Trim$(grdBayDetails(Index).Columns(GRD1_COL_DS).Value) <> "C" _
             And Trim$(grdBayDetails(Index).Columns(GRD1_COL_DS).Value) <> "E" Then
                
             Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, LoadResString(sInvalidDestinationSort), 2250, _
                                              grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                              grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Left, _
                                              grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Width, _
                                              grdBayDetails(Index).RowHeight)
             Set oLoadRoutings = Nothing
             If g_iDebug = 13 Then
                 Call InfoLog("frmPDD PerformFieldLevelValidation - End - ElseIf Trim$(grdBayDetails(Index).Columns(GRD1_COL_DS).Value) <> 'C'...: GRD1_COL_DS = " & GRD1_COL_DS & " - Index: " & Index)
             End If
             Exit Function
          End If
           
          '/////////////////////////////////////////////////////////////////////////////////////////////
          'The value of the Destination Sort determines if we validate the Destination Site with TSLIC./
          '/////////////////////////////////////////////////////////////////////////////////////////////
         
          
          If Trim$(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Value) = HFCS_SORT_TYPE_CODE_CPU Then
              'We do not have to validate the Origin site further
             
          Else
              sDestSite = Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value)
              If Trim$(sDestSite) = "EMPTY" Or ValidateSLIC(sDestSite, sLocalCn, sSlic, sCountryCode) Then
                If sCountryCode = sLocalCn Then
                  grdBayDetails(Index).Columns(GRD1_COL_DEST).Value = sSlic
                End If
              Else
                If grdBayDetails(Index).Visible Then
                  grdBayDetails(Index).Enabled = True
                  grdBayDetails(Index).SetFocus
                End If
 
                grdBayDetails(Index).Columns(GRD1_COL_DS).Value = "C"
              End If
          End If
         
          If grdBayDetails(Index).Columns(GRD1_COL_DS).DataChanged = True Then
              grdBayDetails(Index).Columns(GRD1_COL_RI).Value = "G"
          End If
    Case GRD1_COL_RI
        If (bSaveData = False Or bPredispatchEmpty(Index)) Then
            If pInst.PDInfo.hPDHll(Index + 1).SameLoadName = False Then bValidateLoadRouting = True
           
            If bValidateLoadRouting = False And Index = 1 Then
               If pInst.PDInfo.hPDHll(2).SameLoadName = True And pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True Then
                    bValidateLoadRouting = True
               End If
            End If
           
            If Index = 2 And bValidateLoadRouting = False Then
                If pInst.PDInfo.hPDHll(3).SameLoadName = True And (pInst.PDInfo.hPDHll(2).InvalidLoadRouting = True Or _
                    pInst.PDInfo.hPDHll(1).InvalidLoadRouting = True) Then
                    bValidateLoadRouting = True
                End If
            End If
           
            bGridGotFocus = False
           
            If bValidateLoadRouting Then
                If Trim$(grdBayDetails(Index).Columns(grdBayDetails(Index).Col).Value) = "" Then
'                    If grdBayDetails(Index).Visible Then
'                        grdBayDetails(Index).SetFocus
'                        If Index >= 1 Then DoEvents
'                    End If
                    bLoadRoutingError = True
                    bGridGotFocus = True
                    Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, "Route Code not entered, enter route code", 2250, _
                                                         grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                                         grdBayDetails(Index).Columns(GRD1_COL_RI).Left, _
                                                         grdBayDetails(Index).Columns(GRD1_COL_RI).Width, _
                                                         grdBayDetails(Index).RowHeight)
                    Set oLoadRoutings = Nothing
                    If g_iDebug = 13 Then
                        Call InfoLog("frmPDD PerformFieldLevelValidation - End -  If bValidateLoadRouting - Index: " & Index)
                    End If
                    Exit Function
                Else
                    bValidRouting = True
                    Set oLoadRoutings.RoutingData = New HFCSLoadRoutingCodes.LoadRoutingData
                    iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, "/")
              
                    If iPosition = 0 Then
                        iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_DEST).Value, " ")
                    End If
                   
                    If iPosition > 0 And Len(Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value)) >= iPosition Then
                       sDestSite = Mid$(Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value), 1, iPosition - 1)
                        sCountryCode = Mid$(Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value), iPosition + 1, 2)
                    Else
                        sDestSite = Trim$(grdBayDetails(Index).Columns(GRD1_COL_DEST).Value)
                        sCountryCode = pInst.PDInfo.szCurrentCny
                    End If
                   
                    bFound = oClsDb.GetSlicData(sCountryCode, sDestSite, sSlic)
                   
                    oLoadRoutings.RoutingData.DestCountry = sCountryCode
                    If bFound = True Then
                        oLoadRoutings.RoutingData.DestLocation = sSlic
                    Else
                        oLoadRoutings.RoutingData.DestLocation = sDestSite
                    End If
                   
                    oLoadRoutings.RoutingData.DestSort = grdBayDetails(Index).Columns(GRD1_COL_DS).Value
                   
                    If Not IsNull(pInst.PDInfo.szPDElocCn) Then
                        If Len(Trim$(pInst.PDInfo.szPDElocCn)) > 0 Then
                            sElocCnyCd = pInst.PDInfo.szPDElocCn
                        Else
                            sElocCnyCd = pInst.PDInfo.szCurrentCny
                        End If
                    Else
                        sElocCnyCd = pInst.PDInfo.szCurrentCny
                    End If
                   
                    bFound = oClsDb.GetSlicData(sElocCnyCd, pInst.PDInfo.szPDEloc, sElocName)
                    If bFound = True Then
                        oLoadRoutings.RoutingData.IntermediateLocation = sElocName
                    Else
                        oLoadRoutings.RoutingData.IntermediateLocation = Trim$(pInst.PDInfo.szPDEloc)
                    End If
                   
                    iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_ORIG).Value, "/")
               
                    If iPosition = 0 Then
                        iPosition = InStr(1, grdBayDetails(Index).Columns(GRD1_COL_ORIG).Value, " ")
                    End If
              
                    If iPosition > 0 And Len(Trim$(grdBayDetails(Index).Columns(GRD1_COL_ORIG))) >= iPosition Then
                        sOriginSite = Mid$(Trim$(grdBayDetails(Index).Columns(GRD1_COL_ORIG).Value), 1, iPosition - 1)
                        sCountryCode = Mid$(Trim$(grdBayDetails(Index).Columns(GRD1_COL_ORIG).Value), iPosition + 1, 2)
                    Else
                        sOriginSite = Trim$(grdBayDetails(Index).Columns(GRD1_COL_ORIG).Value)
                        sCountryCode = pInst.PDInfo.szCurrentCny
                    End If
                    
                    bFound = oClsDb.GetSlicData(sCountryCode, sOriginSite, sOrigSlic)
                   
                    If bFound = True Then
                        oLoadRoutings.RoutingData.OrigLocation = sOrigSlic
                    Else
                        oLoadRoutings.RoutingData.OrigLocation = sOriginSite
                    End If
                   
                    oLoadRoutings.RoutingData.OrigCountry = sCountryCode
                    oLoadRoutings.RoutingData.OrigSort = grdBayDetails(Index).Columns(GRD1_COL_OS).Value
                   
                    g_sRouteCodes = oLoadRoutings.getRoutingCodes
                    sOriginalRouteCd = Trim$(grdBayDetails(Index).Columns(GRD1_COL_RI).Value)
                    If Mid$(oLoadRoutings.RoutingCodeStatus, 1, 19) <> "Routing codes found" Then
                        If g_iDebug = 13 Then
                            Call InfoLog("frmPDD PerformFieldLevelValidation - no routing codes found - Index: " & Index)
                        End If
                        bLoadRoutingError = True
                        If Mid$(oLoadRoutings.RoutingCodeStatus, 1, 22) = "No Load Routings found" Then
                            iRc = MsgBox(oLoadRoutings.RoutingCodeStatus & ", do you want to continue?", vbYesNo)
                            If iRc = vbYes Then
                                bValidRouting = False
                                pInst.PDInfo.hPDHll(Index + 1).InvalidLoadRouting = True
                               
'                                If bPredispatchEmpty(Index) = True And bSaveData = True Then 'DoEvents
                                    '- Removed 9-16-14 to resolve flow issues with the screen
                            Else
                                PerformFieldLevelValidation = False
                                bBayCompeleted(Index) = False
                                bBalNrValidated(Index) = False
                                bPredispatchEmpty(Index) = False
                                bSetFocusonBay = True
                                ctlBayNr(Index).Enabled = True
                                ctlBayNr(Index).Visible = True
                                If bSaveData = False Then
                                    If Len(ctlBayNr(Index).Text) = 0 Then
                                        grdBayDetails(Index).Split = 0
                                        grdBayDetails(Index).Col = 0
                                        ctlBayNr(Index).SetFocus
                                    End If
                                End If
                                Set oLoadRoutings = Nothing
                                If g_iDebug = 13 Then
                                    Call InfoLog("frmPDD PerformFieldLevelValidation - End -  If Mid$(oLoadRoutings.RoutingCodeStatus, 1, 22) = No Load Routings found - Index: " & Index)
                                End If
                                Exit Function
                            End If
                        Else
                            Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, oLoadRoutings.RoutingCodeStatus, 2250, _
                                                             grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                                             grdBayDetails(Index).Columns(GRD1_COL_RI).Left, _
                                                             grdBayDetails(Index).Columns(GRD1_COL_RI).Width, _
                                                             grdBayDetails(Index).RowHeight)
                        End If
                    ElseIf IsEmpty(g_sRouteCodes) Then
                        bLoadRoutingError = True
                        If g_iDebug = 13 Then
                            Call InfoLog("frmPDD PerformFieldLevelValidation - empty g_sRouteCodes - Index: " & Index)
                        End If
                        If Mid$(oLoadRoutings.RoutingCodeStatus, 1, 22) = "No Load Routings found" Then
                            iRc = MsgBox(oLoadRoutings.RoutingCodeStatus & ", do you want to continue?", vbYesNo)
                            If iRc = vbYes Then
                                bValidRouting = False
                                pInst.PDInfo.hPDHll(Index + 1).InvalidLoadRouting = True
'                                If bPredispatchEmpty(Index) = True And bSaveData = True Then 'DoEvents - Removed 9-16-14 to resolve flow issues with the screents
                            Else
                                PerformFieldLevelValidation = False
                                bBayCompeleted(Index) = False
                                bBalNrValidated(Index) = False
                                bPredispatchEmpty(Index) = False
                                bSetFocusonBay = True
                                ctlBayNr(Index).Enabled = True
                                ctlBayNr(Index).Visible = True
                                If bSaveData = False Then
                                    If Len(ctlBayNr(Index).Text) = 0 Then
                                        grdBayDetails(Index).Split = 0
                                        grdBayDetails(Index).Col = 0
                                        ctlBayNr(Index).SetFocus
                                    End If
                                End If
                                Set oLoadRoutings = Nothing
                                If g_iDebug = 13 Then
                                    Call InfoLog("frmPDD PerformFieldLevelValidation - End -  ElseIf IsEmpty(g_sRouteCodes) - Index: " & Index)
                                End If
                                Exit Function
                            End If
                        Else
                            Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, "Load Routings not available for load and ELOC.", 2250, _
                                                             grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                                             grdBayDetails(Index).Columns(GRD1_COL_RI).Left, _
                                                             grdBayDetails(Index).Columns(GRD1_COL_RI).Width, _
                                                             grdBayDetails(Index).RowHeight)
                        End If
                    ElseIf Mid$(oLoadRoutings.RoutingCodeStatus, 1, 19) = "Routing codes found" Then
                        'check to see if the route code in the cell is valid
                        For iCounter = 0 To UBound(g_sRouteCodes)
                            If Trim$(g_sRouteCodes(iCounter)) = Trim$(grdBayDetails(Index).Columns(GRD1_COL_RI).Value) Then
                                bValidRouting = True
                                Exit For
                            Else
                                bValidRouting = False
                            End If
                        Next iCounter
                       
                        If bValidRouting = False Then
                            'populate load routings form with Load Routings
                            g_bRoutingOpen = True
                           
                            g_sLoadRoutings = oLoadRoutings.ReturnLoadRoutings
                            g_lTransitDays = oLoadRoutings.TransitDays
                           
                            frmLoadRoutings.Show vbModal, frmPDD
                           
                            Do While g_bRoutingOpen
                                DoEvents
                            Loop
                           
                            If sOriginalRouteCd <> Trim$(grdBayDetails(Index).Columns(GRD1_COL_RI).Value) Then
                                'check again, just in case user did not choose a load routing from the load routing form
                                For iCounter = 0 To UBound(g_sRouteCodes)
                                    If Trim$(g_sRouteCodes(iCounter)) = Trim$(grdBayDetails(Index).Columns(GRD1_COL_RI).Value) Then
                                        bValidRouting = True
                                        Exit For
                                    Else
                                        bValidRouting = False
                                    End If
                                Next iCounter
                            End If
                        End If
                       
                        If bValidRouting = False Then
                            pInst.PDInfo.hPDHll(Index + 1).InvalidLoadRouting = True
                        Else
                            pInst.PDInfo.hPDHll(Index + 1).InvalidLoadRouting = False
                        End If
                    End If
       
                    Set oLoadRoutings.RoutingData = Nothing
                End If
            End If
        End If
    Case GRD1_COL_TYPE
        If bPredispatchEmpty(Index) = True Then
           Set rsTrailerTypes = oClsDb.Get_Valid_Trailer_Type_Codes
          
           iPosition = InStr(1, Trim$(grdBayDetails(Index).Columns(GRD1_COL_TYPE).Value), "*")
          
           If iPosition > 0 Then
                sTlrTyp = Mid$(Trim$(grdBayDetails(Index).Columns(GRD1_COL_TYPE).Value), 1, iPosition - 1)
           Else
                sTlrTyp = Trim$(grdBayDetails(Index).Columns(GRD1_COL_TYPE).Value)
           End If
          
           If Not rsTrailerTypes Is Nothing Then
                If rsTrailerTypes.RecordCount > 0 Then
                    rsTrailerTypes.MoveFirst
                    rsTrailerTypes.Find "CD = '" & sTlrTyp & "' "
                    If Not rsTrailerTypes.EOF Then
                        rsTrailerTypes.MoveFirst
                        bTrailerError = False
                    Else
                        rsTrailerTypes.MoveFirst
                        bTrailerError = True
                       'have to force the grid to set focus, otherwise, it will always go to the dolly field.
                        If grdBayDetails(Index).Visible Then
                            grdBayDetails(Index).Enabled = True
                            grdBayDetails(Index).SetFocus
                        End If
                        Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, "Trailer Type not valid.", 2250, _
                                     grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                     grdBayDetails(Index).Columns(GRD1_COL_TYPE).Left, _
                                     grdBayDetails(Index).Columns(GRD1_COL_TYPE).Width, _
                                     grdBayDetails(Index).RowHeight)
 
                        Call oClsDb.DBClass.CloseRecordSet(rsTrailerTypes)
                        Set oLoadRoutings = Nothing
                        If g_iDebug = 13 Then
                            Call InfoLog("frmPDD PerformFieldLevelValidation - End -  If Not rsTrailerTypes Is Nothing - Index: " & Index)
                        End If
                        Exit Function
                    End If
                Else
                    'have to force the grid to set focus, otherwise, it will always go to the dolly field.
                    If grdBayDetails(Index).Visible Then
                        grdBayDetails(Index).Enabled = True
                        grdBayDetails(Index).SetFocus
                    End If
                    bTrailerError = True
                    Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, "Trailer Type not valid.", 2250, _
                                 grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                                 grdBayDetails(Index).Columns(GRD1_COL_TYPE).Left, _
                                 grdBayDetails(Index).Columns(GRD1_COL_TYPE).Width, _
                                 grdBayDetails(Index).RowHeight)
 
                    Call oClsDb.DBClass.CloseRecordSet(rsTrailerTypes)
                    Set oLoadRoutings = Nothing
                    If g_iDebug = 13 Then
                        Call InfoLog("frmPDD PerformFieldLevelValidation - End -  ELSE NOT rsTrailerTypes.RecordCount > 0 - Index: " & Index)
                    End If
                    Exit Function
                End If
           Else
                'have to force the grid to set focus, otherwise, it will always go to the dolly field.
                If grdBayDetails(Index).Visible Then
                    grdBayDetails(Index).Enabled = True
                    grdBayDetails(Index).SetFocus
                End If
                bTrailerError = True
                Call ctlWarningPopup.ShowWarning(frmPDD.grdBayDetails(Index).hwnd, "Trailer Type not valid.", 2250, _
                             grdBayDetails(Index).RowTop(grdBayDetails(Index).Row), _
                             grdBayDetails(Index).Columns(GRD1_COL_TYPE).Left, _
                             grdBayDetails(Index).Columns(GRD1_COL_TYPE).Width, _
                             grdBayDetails(Index).RowHeight)
 
                Call oClsDb.DBClass.CloseRecordSet(rsTrailerTypes)
                Set oLoadRoutings = Nothing
                If g_iDebug = 13 Then
                    Call InfoLog("frmPDD PerformFieldLevelValidation - End - ELSE Not rsTrailerTypes Is Nothing - Index: " & Index)
                End If
                Exit Function
           End If
          
           Call oClsDb.DBClass.CloseRecordSet(rsTrailerTypes)
        End If
    Case Else
  End Select
      Set oLoadRoutings = Nothing
 
  'If we made it this far then we are good.
  PerformFieldLevelValidation = True
  Set oLoadRoutings = Nothing
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD PerformFieldLevelValidation - End - Index: " & Index)
    End If
    Exit Function
Error_Handler:
    glErrNum = 400
    gsError = "PerformFieldLevelValidation"
    Call oProc.update_error_object(Me, gsError)
    Set oLoadRoutings = Nothing
 
    Screen.MousePointer = vbDefault
     If g_iDebug = 13 Then
       Call InfoLog("frmPDD PerformFieldLevelValidation error " & Err.Description)
    End If
    If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module: PerformFieldLevelValidation", _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
    '    Unload Me
    End If
End Function
 
Private Function ValidateSLIC(ByVal sSlicCn As String, _
                                 ByVal sLocalCn As String, _
                                 ByRef sSlic As String, _
                                 ByRef sCountry As String) As Boolean
    Dim iPosition As Integer
   
    On Error GoTo ErrorHandler
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ValidateSLIC - Begin")
    End If
     ' field can not be empty
    If Len(Trim$(sSlicCn)) = 0 Then
        ValidateSLIC = False
        If g_iDebug = 13 Then
           Call InfoLog("frmPDD ValidateSLIC - END - ssliccn")
        End If
        Exit Function
    End If
   
     SplitSabCny sSlicCn, sSlic, sCountry, pInst.PDInfo.szCurrentCny
  
    
    ValidateSLIC = F9802VerifySLC_CN(sSlic, sCountry)
    If g_iDebug = 13 Then
       Call InfoLog("frmPDD ValidateSLIC - End")
    End If
    Exit Function
ErrorHandler:
glErrNum = 400
gsError = "ValidateSLIC"
Call oProc.update_error_object(Me, gsError)
 
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                            IIf(Err.Number <> 0, Err.Number, glErrNum), _
                            Err.Description & " Module:" & gsError, _
                            oProc, _
                            ERROR_MSG, _
                            FEEDER_DISPATCH_DRIVER) Then
Set oEventlog = Nothing
MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
' Unload Me
End If
End Function
 
Private Sub SetGridForEdit(Index As Integer)
Dim CustomCellStyle As TrueOleDBGrid70.Style
On Error GoTo Error_Handler
If g_iDebug = 13 Then
    Call InfoLog("frmPDD SetGridForEdit - Begin: Index = " & Index)
End If
 
 
If bPredispatchEmpty(Index) Or bPredispatchEmptyWithTrailer(Index) Then
    With grdBayDetails(Index)
      .TabAction = dbgColumnNavigation
      .TabAcrossSplits = True
      .DirectionAfterEnter = dbgMoveRight
      .Splits(2).SelectedStyle.Font.Bold = True
      .Splits(2).SelectedBackColor = vbGreen
      .Splits(2).SelectedStyle.BackColor = vbGreen
      .Splits(2).HighlightRowStyle.BackColor = vbGreen
      .Splits(2).MarqueeStyle = dbgHighlightCell '+ dbgFloatingEditor 'dbgHighlightRowRaiseCell
     
      .Splits(2).Columns(GRD1_COL_DEST).AllowFocus = True
      .Splits(2).Columns(GRD1_COL_DEST).Locked = False
      .Splits(2).Columns(GRD1_COL_DEST).EditBackColor = vbGreen
      .Splits(2).Columns(GRD1_COL_DEST).EditMask = "?&&&&&??"
      .Splits(2).Columns(GRD1_COL_DEST).DataWidth = 8
     
      .Splits(2).Columns(GRD1_COL_DS).AllowFocus = True
      .Splits(2).Columns(GRD1_COL_DS).Locked = False
      .Splits(2).Columns(GRD1_COL_DS).EditBackColor = vbGreen
      .Splits(2).Columns(GRD1_COL_DS).EditMask = ">@"
      .Splits(2).Columns(GRD1_COL_DS).DataWidth = 1
     
      .Splits(3).SelectedStyle.Font.Bold = True
      .Splits(3).SelectedBackColor = vbGreen
      .Splits(3).SelectedStyle.BackColor = vbGreen
      .Splits(3).HighlightRowStyle.BackColor = vbGreen
      .Splits(3).MarqueeStyle = dbgHighlightCell '+ dbgFloatingEditor ' dbgHighlightRowRaiseCell 'dbgFloatingEditor
   
      .Splits(3).Columns(GRD1_COL_RI).AllowFocus = True
      .Splits(3).Columns(GRD1_COL_RI).Locked = False
      .Splits(3).Columns(GRD1_COL_RI).EditBackColor = vbGreen
      .Splits(3).Columns(GRD1_COL_RI).EditMask = ">@?"
      .Splits(3).Columns(GRD1_COL_RI).DataWidth = 2
       
      .Splits(1).SelectedStyle.Font.Bold = True
      .Splits(1).SelectedBackColor = vbGreen
      .Splits(1).SelectedStyle.BackColor = vbGreen
      .Splits(1).HighlightRowStyle.BackColor = vbGreen
      .Splits(1).MarqueeStyle = dbgHighlightCell '+ dbgFloatingEditor
       
      .Splits(1).Columns(GRD1_COL_TYPE).AllowFocus = True
      .Splits(1).Columns(GRD1_COL_TYPE).Locked = False
      .Splits(1).Columns(GRD1_COL_TYPE).EditBackColor = vbGreen
      .Splits(1).Columns(GRD1_COL_TYPE).DataWidth = 4
    End With
ElseIf bPredispatchEmpty(Index) = False And bPredispatchEmptyWithTrailer(Index) = False Then
    With grdBayDetails(Index)
      .TabAction = dbgColumnNavigation
      .TabAcrossSplits = True
      .DirectionAfterEnter = dbgMoveRight
      .Splits(3).SelectedStyle.Font.Bold = True
      .Splits(3).SelectedBackColor = vbGreen
      .Splits(3).SelectedStyle.BackColor = vbGreen
      .Splits(3).HighlightRowStyle.BackColor = vbGreen
      .Splits(3).MarqueeStyle = dbgHighlightCell '+ dbgFloatingEditor ' dbgHighlightRowRaiseCell 'dbgFloatingEditor
      .Splits(3).Columns(GRD1_COL_RI).AllowFocus = True
      .Splits(3).Columns(GRD1_COL_RI).Locked = False
      .Splits(3).Columns(GRD1_COL_RI).EditBackColor = vbGreen
      .Splits(3).Columns(GRD1_COL_RI).EditMask = ">@?"
      .Splits(3).Columns(GRD1_COL_RI).DataWidth = 2
    End With
End If
 
If g_iDebug = 13 Then
    Call InfoLog("frmPDD SetGridForEdit - End: Index = " & Index)
End If
Exit Sub
 
Error_Handler:
glErrNum = 400
gsError = "SetGridForEdit"
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
     '   Unload Me
End If
End Sub
 
Private Sub SetGridNormal(Index As Integer)
On Error GoTo Error_Handler
 
If g_iDebug = 13 Then
    Call InfoLog("frmPDD SetGridNormal - Begin: Index = " & Index)
End If
 
If bPredispatchEmpty(Index) Or bPredispatchEmptyWithTrailer(Index) Then
  With grdBayDetails(Index)
      .Splits(2).SelectedStyle.Font.Bold = False
      .Splits(2).SelectedBackColor = vbWhite
      .Splits(2).SelectedStyle.BackColor = vbWhite
      .Splits(2).HighlightRowStyle.BackColor = vbWhite
      .Splits(2).MarqueeStyle = dbgNoMarquee
     
      .Splits(2).Columns(GRD1_COL_DEST).BackColor = vbWhite
      .Splits(2).Columns(GRD1_COL_DEST).Font.Bold = False
      .Splits(2).Columns(GRD1_COL_DEST).AllowFocus = False
    '  .Splits(2).Columns(GRD1_COL_DEST).Locked = True
     
      .Splits(2).Columns(GRD1_COL_DS).BackColor = vbWhite
      .Splits(2).Columns(GRD1_COL_DS).Font.Bold = False
      .Splits(2).Columns(GRD1_COL_DS).AllowFocus = False
     ' .Splits(2).Columns(GRD1_COL_DS).Locked = True
     
      .Splits(3).SelectedStyle.Font.Bold = False
      .Splits(3).SelectedBackColor = vbWhite
      .Splits(3).SelectedStyle.BackColor = vbWhite
      .Splits(3).HighlightRowStyle.BackColor = vbWhite
      .Splits(3).MarqueeStyle = dbgNoMarquee
     
      .Splits(3).Columns(GRD1_COL_RI).BackColor = vbWhite
      .Splits(3).Columns(GRD1_COL_RI).Font.Bold = False
      .Splits(3).Columns(GRD1_COL_RI).AllowFocus = False
    '  .Splits(3).Columns(GRD1_COL_RI).Locked = True
     
        
      If Len(Trim$(ctlBayNr(Index).Text)) = 0 Then
          .Splits(1).SelectedStyle.Font.Bold = False
          .Splits(1).SelectedBackColor = vbWhite
          .Splits(1).SelectedStyle.BackColor = vbWhite
          .Splits(1).HighlightRowStyle.BackColor = vbWhite
          .Splits(1).MarqueeStyle = dbgNoMarquee
        
          .Splits(1).Columns(GRD1_COL_TYPE).BackColor = vbWhite
          .Splits(1).Columns(GRD1_COL_TYPE).Font.Bold = False
          .Splits(1).Columns(GRD1_COL_TYPE).AllowFocus = False
         
      End If
  End With
Else
  grdBayDetails(Index).Splits(3).SelectedStyle.Font.Bold = False
  grdBayDetails(Index).Splits(3).SelectedBackColor = vbWhite
  grdBayDetails(Index).Splits(3).SelectedStyle.BackColor = vbWhite
  grdBayDetails(Index).Splits(3).HighlightRowStyle.BackColor = vbWhite
  grdBayDetails(Index).Splits(3).MarqueeStyle = dbgNoMarquee
 
  grdBayDetails(Index).Splits(3).Columns(GRD1_COL_RI).BackColor = vbWhite
  grdBayDetails(Index).Splits(3).Columns(GRD1_COL_RI).Font.Bold = False
End If
 
If g_iDebug = 13 Then
    Call InfoLog("frmPDD SetGridNormal - End: Index = " & Index)
End If
Exit Sub
Error_Handler:
glErrNum = 400
gsError = "SetGridNormal"
Call oProc.update_error_object(Me, gsError)
  Screen.MousePointer = vbDefault
  If oErrorObject.error_routine(oEventlog.FeederShell, _
                                   IIf(Err.Number <> 0, Err.Number, glErrNum), _
                                   Err.Description & " Module:" & gsError, _
                                   oProc, _
                                   ERROR_MSG, _
                                   FEEDER_DISPATCH_DRIVER) Then
        Set oEventlog = Nothing
        MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
   '     Unload Me
End If
End Sub
 
Private Sub PopulateMultiLegGrid()
    Dim i As Integer
    Dim sPrevSegOid As String
    Dim sSegOids As String
    Dim iFirstLegNotDeparted As Integer
    Dim bLegNotDepartedFound As Boolean
    Dim bAddMondaySch As Boolean
   
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD PopulateMultiLegGrid - Begin")
    End If
   
    On Error GoTo ErrorHandler
   
    Set xaMultiLegArray = New XArrayDB
   
    Set rsMultiLeg = oClsDb.GetJobInformation(pInst.PDInfo.szPDJobNr, frmPDD.ctlWkendingDate.Value, CStr(frmPDD.cboDow.ListIndex), , pInst.PDInfo.szPDJobDomCn, pInst.PDInfo.szPDJobDomSlic)
    If rsMultiLeg.RecordCount > 0 Then
        If Not IsNull(rsMultiLeg.Fields("JOB_SYS_NR_OID").Value) Then
            JobSysNrOid = rsMultiLeg.Fields("JOB_SYS_NR_OID").Value
        End If
    End If
   
    Set rsMultiLeg = oClsDb.GetLegInformation(pInst.PDInfo.szPDJobNr, frmPDD.ctlWkendingDate.Value, CStr(frmPDD.cboDow.ListIndex), , pInst.PDInfo.szPDJobDomCn, pInst.PDInfo.szPDJobDomSlic)
   
    sPrevSegOid = ""
   
    Call xaMultiLegArray.ReDim(0, rsMultiLeg.RecordCount, 0, 9)
   
    iFirstLegNotDeparted = 0
    bLegNotDepartedFound = False
   
    If rsMultiLeg.RecordCount > 0 Then
        rsMultiLeg.MoveFirst
        bAddMondaySch = True
        'TT00751 and TT00752
        If szFdrSchDow = "0" Then
            If Not IsNull(rsMultiLeg.Fields("fdr_sch_dow_cd").Value) Then
                If rsMultiLeg.Fields("fdr_sch_dow_cd").Value = szFdrSchDow Then
                  bAddMondaySch = False
                Else
                  bAddMondaySch = True
                End If
            Else
                bAddMondaySch = True
            End If
        End If
    End If
   
    Do While Not rsMultiLeg.EOF
        If sPrevSegOid <> Trim$(rsMultiLeg.Fields("SEG_SYS_NR_OID").Value) And (rsMultiLeg.Fields("fdr_sch_dow_cd").Value = szFdrSchDow Or bAddMondaySch = True) Then
            xaMultiLegArray.Value(i, 0) = Format$(rsMultiLeg.Fields("fdr_acy_sch_stt_dt").Value, "mm/dd") & " " & Format$(rsMultiLeg.Fields("fdr_acy_sch_stt_tm").Value, "hh:mm")
            xaMultiLegArray.Value(i, 1) = rsMultiLeg.Fields("SCH_MVM_DTN_SAB_NA").Value
            xaMultiLegArray.Value(i, 2) = rsMultiLeg.Fields("SEG_SYS_NR_OID").Value
            xaMultiLegArray.Value(i, 3) = rsMultiLeg.Fields("DVR_LST_NA").Value
            xaMultiLegArray.Value(i, 4) = rsMultiLeg.Fields("FDR_JOB_OWN_SAB_NA").Value
            xaMultiLegArray.Value(i, 5) = rsMultiLeg.Fields("FDR_JOB_OWN_CNY_CD").Value
            xaMultiLegArray.Value(i, 6) = rsMultiLeg.Fields("SCH_TRC_EQP_NR").Value
            xaMultiLegArray.Value(i, 7) = rsMultiLeg.Fields("SCH_MVM_DTN_CNY_CD").Value
            Set m_rsLegLoadInfo = oClsDb.GetLoadsonLeg(rsMultiLeg.Fields("SEG_SYS_NR_OID").Value, True)
            If m_rsLegLoadInfo.RecordCount > 0 Then
                If IsNull(m_rsLegLoadInfo.Fields("bay_eqp_dpt_ts").Value) Then
                    'check TFEMSGS to see if leg has departed, but loads do not show auto-departed in HFCS
                    bLegNotDepartedFound = Not oClsDb.HasLegDeparted(rsMultiLeg.Fields("SEG_SYS_NR_OID").Value)
                    If bLegNotDepartedFound = True Then
                        xaMultiLegArray.Value(i, 8) = False
                    Else
                        xaMultiLegArray.Value(i, 8) = True
                        iFirstLegNotDeparted = iFirstLegNotDeparted + 1
                    End If
                    'bLegNotDepartedFound = True
                Else
                    xaMultiLegArray.Value(i, 8) = True
                    If bLegNotDepartedFound = False Then
                        iFirstLegNotDeparted = iFirstLegNotDeparted + 1
                    End If
                End If
            Else
                'check TFEMSGS to see if leg has departed, but loads do not show auto-departed in HFCS
                bLegNotDepartedFound = Not oClsDb.HasLegDeparted(rsMultiLeg.Fields("SEG_SYS_NR_OID").Value)
                If bLegNotDepartedFound = True Then
                    xaMultiLegArray.Value(i, 8) = False
                Else
                    xaMultiLegArray.Value(i, 8) = True
                    If bLegNotDepartedFound = False Then
                        iFirstLegNotDeparted = iFirstLegNotDeparted + 1
                    End If
                End If
            End If
            xaMultiLegArray.Value(i, 9) = rsMultiLeg.Fields("JOB_SYS_NR_OID").Value
            sSegOids = "'" & rsMultiLeg.Fields("SEG_SYS_NR_OID").Value & "',"
            i = i + 1
        End If
        sPrevSegOid = rsMultiLeg.Fields("SEG_SYS_NR_OID").Value
        rsMultiLeg.MoveNext
    Loop
   
    If Len(Trim$(sSegOids)) > 0 Then
        sSegOids = Mid$(sSegOids, 1, Len(sSegOids) - 1)
    Else
        MsgBox "There are no scheduled legs for selected job."
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD PopulateMultiLegGrid - End - There are no scheduled legs for selected job.")
        End If
        Exit Sub
    End If
    Call xaMultiLegArray.ReDim(0, i - 1, 0, 9)
    grdMultLegView.Array = xaMultiLegArray
   
    grdMultLegView.ReBind
    If xaMultiLegArray.Value(0, 8) = True Then
        If i > 1 Then
            grdMultLegView.Bookmark = iFirstLegNotDeparted 'grdMultLegView.Bookmark + 1
        End If
    End If
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD PopulateMultiLegGrid - End")
    End If
Exit Sub
 
ErrorHandler:
    glErrNum = 400
    gsError = "PopulateMultiLegGrid"
    'the object has disconnect from its clients grid error, retry the grid rebind
    If Err.Number = -2147417848 Then
        If g_iDebug = 13 Then
            Call InfoLog("frmPDD ClearField - received error -2147417848")
        End If
         
        If RecoverGridError(grdMultLegView) = True Then
            Resume Next
        End If
    End If
    Call oProc.update_error_object(Me, gsError)
   With Err
      'Handle the errror here as you normally would
      Call oErrorObject.error_routine(oEventlog.GlobalEvents, _
                                    .Number, .Description & " Module:" & gsError, _
                                    oProc, ERROR_MSG, GLOBAL_LOADOBJECT, _
                                    NO_POP_UP)
      Call .Raise(LOAD_ERRORS.LOADS_OBJ_UNKNOWN_ERROR, .Source, .Description, .HelpFile, .HelpContext)
  End With
  MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
' Unload Me
End Sub
 
Private Sub function_key_pressed(KeyCode As Integer, Shift As Integer)
    On Error GoTo Error_Handler
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD function_key_pressed - Begin")
    End If
    Select Case KeyCode
        Case oKeyDefs.Enter_Key
            If (rsMultiLeg.RecordCount > 0) Then
                pInst.PDInfo.szPDJobDomSlic = xaMultiLegArray.Value(grdMultLegView.Bookmark, 4)
                pInst.PDInfo.szPDJobDomCn = xaMultiLegArray.Value(grdMultLegView.Bookmark, 5)
                If Not IsNull(xaMultiLegArray.Value(grdMultLegView.Bookmark, 3)) Then
                    pInst.PDInfo.szPDDvrNa = Trim$(xaMultiLegArray.Value(grdMultLegView.Bookmark, 3))
                Else
                    pInst.PDInfo.szPDDvrNa = ""
                End If
               
                If Not IsNull(xaMultiLegArray.Value(grdMultLegView.Bookmark, 6)) Then
                    pInst.PDInfo.szPDTrcNr = xaMultiLegArray.Value(grdMultLegView.Bookmark, 6)
                Else
                    pInst.PDInfo.szPDTrcNr = Trim$(txtTractorNumber.Text)
                End If
                pInst.PDInfo.szPDEloc = xaMultiLegArray.Value(grdMultLegView.Bookmark, 1)
                pInst.PDInfo.SegmentSystemNumberOID = xaMultiLegArray.Value(grdMultLegView.Bookmark, 2)
                pInst.PDInfo.szPDElocCn = xaMultiLegArray.Value(grdMultLegView.Bookmark, 7)
                pInst.PDInfo.FeederScheduleEndTime = xaMultiLegArray.Value(grdMultLegView.Bookmark, 0)
                pInst.PDInfo.szFdrSchWndDT = Format$(frmPDD.ctlWkendingDate.Value, "mm/dd/yyyy")
                pInst.PDInfo.JobSystemNumberOID = xaMultiLegArray.Value(grdMultLegView.Bookmark, 9)
            End If
    End Select
    If g_iDebug = 13 Then
        Call InfoLog("frmPDD function_key_pressed - End")
    End If
    Exit Sub
Error_Handler:
glErrNum = 400
gsError = "Function_key_pressed"
Call oProc.update_error_object(Me, gsError)
Screen.MousePointer = vbDefault
If oErrorObject.error_routine(oEventlog.FeederShell, _
                               IIf(Err.Number <> 0, Err.Number, glErrNum), _
                               Err.Description & " Module:" & gsError, _
                               oProc, _
                               ERROR_MSG, _
                              FEEDER_DISPATCH_DRIVER) Then
    Set oEventlog = Nothing
    MsgBox "An Error occured in <" & gsError & "> Try Restarting the Feeder Shell to Resolve this!!!", vbCritical, "HFCS Error Occured"
'   Unload Me
End If
End Sub
 
Private Function RecoverGridError(grdRecover As TrueOleDBGrid70.TDBGrid) As Boolean
    On Error GoTo Error_Handler
    Me.Refresh
    lRecoverQueries = lRecoverQueries + 1
    Debug.Print "Recover attempt #" & CStr(lRecoverQueries)
    If lRecoverQueries > MAX_RECOVER_QUERIES Then
      RecoverGridError = False
    Else
      Call grdRecover.Refresh
      RecoverGridError = True
    End If
Exit Function
Error_Handler:
    Call InfoLog("RecoverGridError Desc = " & Err.Description & ", num = " & Err.Number, ERROR_MSG)
    Resume Next
End Function
