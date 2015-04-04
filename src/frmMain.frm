VERSION 5.00
Object = "{E7BB2F30-C5DD-4370-B7E2-19A7EDF169EE}#1.3#0"; "TabStripCtlU.ocx"
Object = "{2AFA7915-463D-4B61-AEB7-41B1236C143E}#1.5#0"; "BtnCtlsU.ocx"
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.3#0"; "TrackBarCtlU.ocx"
Object = "{52D76F35-4551-4C56-B53B-A343E42B0AF8}#2.1#0"; "ProgBarU.ocx"
Begin VB.Form frmMain 
   Caption         =   "Hot Spot Studio by RL Vision"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRestoreLinkLabels 
      Interval        =   50
      Left            =   8160
      Top             =   120
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Additive"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   6840
      Value           =   1  'Checked
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Restore"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   7200
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Snapshot"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   7200
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   3120
      ScaleHeight     =   265
      ScaleMode       =   0  'User
      ScaleWidth      =   246
      TabIndex        =   0
      Top             =   480
      Width           =   3600
      Begin VB.Label lblStartupTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right click here to add a light!"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   39
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Label lblShowIntensity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblShowIntensity"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   38
         Top             =   120
         Width           =   1140
         Visible         =   0   'False
      End
      Begin VB.Shape shpLight 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         Shape           =   3  'Circle
         Top             =   120
         Width           =   495
         Visible         =   0   'False
      End
   End
   Begin VB.PictureBox picRender 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   3120
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   4
      Top             =   480
      Width           =   3555
   End
   Begin TabStripCtlLibUCtl.TabStrip TabStrip1 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      _cx             =   4895
      _cy             =   10398
      AllowDragDrop   =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      CloseableTabs   =   0   'False
      DisabledEvents  =   262379
      DragActivateTime=   -1
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      FixedTabWidth   =   96
      FocusOnButtonDown=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HorizontalPadding=   6
      HotTracking     =   0   'False
      HoverTime       =   -1
      InsertMarkColor =   0
      MinTabWidth     =   -1
      MousePointer    =   0
      MultiRow        =   0   'False
      MultiSelect     =   0   'False
      OLEDragImageStyle=   0
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      RaggedTabRows   =   -1  'True
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ScrollToOpposite=   0   'False
      ShowButtonSeparators=   -1  'True
      ShowToolTips    =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TabBoundingBoxDefinition=   4
      TabCaptionAlignment=   0
      TabHeight       =   18
      TabPlacement    =   0
      UseFixedTabWidth=   0   'False
      UseSystemFont   =   -1  'True
      VerticalPadding =   3
      Begin BtnCtlsLibUCtl.Frame frmHelp 
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   2535
         _cx             =   4471
         _cy             =   1085
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         BorderVisible   =   0   'False
         ContentType     =   0
         DisabledEvents  =   3
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   0
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         MousePointer    =   0
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         IconIndexes     =   "frmMain.frx":000C
         Text            =   "frmMain.frx":0044
         Begin VB.Image Image4 
            Height          =   330
            Left            =   0
            Picture         =   "frmMain.frx":0064
            Top             =   135
            Width           =   315
         End
         Begin VB.Image Image5 
            Height          =   345
            Left            =   1440
            Picture         =   "frmMain.frx":04A8
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add / Del"
            Height          =   195
            Left            =   1800
            TabIndex        =   34
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select / Move"
            Height          =   195
            Left            =   360
            TabIndex        =   33
            Top             =   240
            Width           =   1020
         End
      End
      Begin BtnCtlsLibUCtl.CommandButton cmdRender 
         Default         =   -1  'True
         Height          =   495
         Left            =   360
         TabIndex        =   31
         ToolTipText     =   "Make a high quality rendering of this scene"
         Top             =   5160
         Width           =   2055
         _cx             =   3625
         _cy             =   873
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         ButtonType      =   0
         ContentType     =   0
         DisabledEvents  =   1289
         DontRedraw      =   0   'False
         DropDownArrowHeight=   0
         DropDownArrowWidth=   15
         DropDownGlyph   =   54
         DropDownOnRight =   -1  'True
         DropDownPushed  =   0   'False
         DropDownStyle   =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   1
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         KeepDropDownArrowAspectRatio=   -1  'True
         MousePointer    =   0
         MultiLine       =   -1  'True
         ProcessContextMenuKeys=   -1  'True
         Pushed          =   0   'False
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         ShowRightsElevationIcon=   0   'False
         ShowSplitLine   =   -1  'True
         Style           =   0
         SupportOLEDragImages=   -1  'True
         TextMarginBottom=   1
         TextMarginLeft  =   1
         TextMarginRight =   1
         TextMarginTop   =   1
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         VAlignment      =   1
         IconIndexes     =   "frmMain.frx":08ED
         CommandLinkNote =   "frmMain.frx":0925
         Text            =   "frmMain.frx":0945
      End
      Begin BtnCtlsLibUCtl.Frame frmTab12 
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   3960
         Width           =   2535
         _cx             =   4471
         _cy             =   1720
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         BorderVisible   =   -1  'True
         ContentType     =   0
         DisabledEvents  =   3
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   0
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         MousePointer    =   0
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         IconIndexes     =   "frmMain.frx":0981
         Text            =   "frmMain.frx":09B9
         Begin BtnCtlsLibUCtl.CommandButton cmdLoadBG 
            Height          =   375
            Left            =   1320
            TabIndex        =   29
            ToolTipText     =   "Load a background picture for the canvas"
            Top             =   360
            Width           =   1095
            _cx             =   1931
            _cy             =   661
            Appearance      =   0
            BackColor       =   -2147483633
            BorderStyle     =   0
            ButtonType      =   0
            ContentType     =   0
            DisabledEvents  =   1289
            DontRedraw      =   0   'False
            DropDownArrowHeight=   0
            DropDownArrowWidth=   15
            DropDownGlyph   =   54
            DropDownOnRight =   -1  'True
            DropDownPushed  =   0   'False
            DropDownStyle   =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            HAlignment      =   1
            HoverTime       =   -1
            IconAlignment   =   0
            IconMarginBottom=   0
            IconMarginLeft  =   0
            IconMarginRight =   0
            IconMarginTop   =   0
            KeepDropDownArrowAspectRatio=   -1  'True
            MousePointer    =   0
            MultiLine       =   -1  'True
            ProcessContextMenuKeys=   -1  'True
            Pushed          =   0   'False
            RegisterForOLEDragDrop=   0   'False
            RightToLeft     =   0
            ShowRightsElevationIcon=   0   'False
            ShowSplitLine   =   -1  'True
            Style           =   0
            SupportOLEDragImages=   -1  'True
            TextMarginBottom=   1
            TextMarginLeft  =   1
            TextMarginRight =   1
            TextMarginTop   =   1
            UseImprovedImageListSupport=   0   'False
            UseSystemFont   =   -1  'True
            VAlignment      =   1
            IconIndexes     =   "frmMain.frx":09E5
            CommandLinkNote =   "frmMain.frx":0A1D
            Text            =   "frmMain.frx":0A3D
         End
         Begin BtnCtlsLibUCtl.CommandButton cmdSetBG 
            Height          =   375
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Set the canvas back colour. The currently selected light colour is used."
            Top             =   360
            Width           =   1095
            _cx             =   1931
            _cy             =   661
            Appearance      =   0
            BackColor       =   -2147483633
            BorderStyle     =   0
            ButtonType      =   0
            ContentType     =   0
            DisabledEvents  =   1289
            DontRedraw      =   0   'False
            DropDownArrowHeight=   0
            DropDownArrowWidth=   15
            DropDownGlyph   =   54
            DropDownOnRight =   -1  'True
            DropDownPushed  =   0   'False
            DropDownStyle   =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            HAlignment      =   1
            HoverTime       =   -1
            IconAlignment   =   0
            IconMarginBottom=   0
            IconMarginLeft  =   0
            IconMarginRight =   0
            IconMarginTop   =   0
            KeepDropDownArrowAspectRatio=   -1  'True
            MousePointer    =   0
            MultiLine       =   -1  'True
            ProcessContextMenuKeys=   -1  'True
            Pushed          =   0   'False
            RegisterForOLEDragDrop=   0   'False
            RightToLeft     =   0
            ShowRightsElevationIcon=   0   'False
            ShowSplitLine   =   -1  'True
            Style           =   0
            SupportOLEDragImages=   -1  'True
            TextMarginBottom=   1
            TextMarginLeft  =   1
            TextMarginRight =   1
            TextMarginTop   =   1
            UseImprovedImageListSupport=   0   'False
            UseSystemFont   =   -1  'True
            VAlignment      =   1
            IconIndexes     =   "frmMain.frx":0A73
            CommandLinkNote =   "frmMain.frx":0AAB
            Text            =   "frmMain.frx":0ACB
         End
      End
      Begin BtnCtlsLibUCtl.Frame frmTab11 
         Height          =   2775
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   2535
         _cx             =   4471
         _cy             =   4895
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         BorderVisible   =   -1  'True
         ContentType     =   0
         DisabledEvents  =   3
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   0
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         MousePointer    =   0
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         IconIndexes     =   "frmMain.frx":0AFF
         Text            =   "frmMain.frx":0B37
         Begin TrackBarCtlLibUCtl.TrackBar TrackBar_Int 
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1920
            Width           =   1935
            _cx             =   3413
            _cy             =   661
            Appearance      =   0
            AutoTickFrequency=   1
            AutoTickMarks   =   -1  'True
            BackColor       =   -2147483633
            BackgroundDrawMode=   0
            BorderStyle     =   0
            CurrentPosition =   2
            DisabledEvents  =   779
            DontRedraw      =   0   'False
            DownIsLeft      =   -1  'True
            Enabled         =   -1  'True
            HoverTime       =   -1
            LargeStepWidth  =   -1
            Maximum         =   20
            Minimum         =   1
            MousePointer    =   0
            Orientation     =   0
            ProcessContextMenuKeys=   -1  'True
            RangeSelectionEnd=   0
            RangeSelectionStart=   0
            RegisterForOLEDragDrop=   0   'False
            Reversed        =   0   'False
            RightToLeftLayout=   0   'False
            SelectionType   =   0
            ShowSlider      =   -1  'True
            SliderLength    =   -1
            SmallStepWidth  =   1
            SupportOLEDragImages=   -1  'True
            TickMarksPosition=   0
            ToolTipPosition =   2
         End
         Begin BtnCtlsLibUCtl.CheckBox Check1 
            Height          =   255
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "Sptolights with this turned on look like a normal spotlights, but will absorb light from spotlights with this turned off."
            Top             =   2400
            Width           =   1815
            _cx             =   3201
            _cy             =   450
            Appearance      =   0
            AutoToggleCheckMark=   -1  'True
            BackColor       =   -2147483633
            BorderStyle     =   0
            CheckMarkOnRight=   0   'False
            ContentType     =   0
            DisabledEvents  =   267
            DontRedraw      =   0   'False
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            HAlignment      =   0
            HoverTime       =   -1
            IconAlignment   =   0
            IconMarginBottom=   0
            IconMarginLeft  =   0
            IconMarginRight =   0
            IconMarginTop   =   0
            MousePointer    =   0
            MultiLine       =   -1  'True
            ProcessContextMenuKeys=   -1  'True
            Pushed          =   0   'False
            PushLike        =   0   'False
            RegisterForOLEDragDrop=   0   'False
            RightToLeft     =   0
            SelectionState  =   0
            Style           =   0
            SupportOLEDragImages=   -1  'True
            TextMarginBottom=   1
            TextMarginLeft  =   1
            TextMarginRight =   1
            TextMarginTop   =   1
            TriState        =   0   'False
            UseImprovedImageListSupport=   0   'False
            UseSystemFont   =   -1  'True
            VAlignment      =   1
            IconIndexes     =   "frmMain.frx":0B73
            Text            =   "frmMain.frx":0BAB
         End
         Begin VB.TextBox txtSize 
            Height          =   285
            Left            =   2040
            TabIndex        =   20
            Tag             =   "tab1"
            Text            =   "50"
            Top             =   1335
            Width           =   375
         End
         Begin VB.TextBox txtInt 
            Height          =   285
            Left            =   2055
            TabIndex        =   19
            Tag             =   "tab1"
            Text            =   "2"
            Top             =   1950
            Width           =   375
         End
         Begin VB.PictureBox picCurrColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00436FFF&
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1935
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   18
            Tag             =   "tab1"
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   255
            Picture         =   "frmMain.frx":0BF3
            ScaleHeight     =   615
            ScaleWidth      =   1575
            TabIndex        =   17
            Tag             =   "tab1"
            Top             =   360
            Width           =   1575
         End
         Begin TrackBarCtlLibUCtl.TrackBar TrackBar_Size 
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   1935
            _cx             =   3413
            _cy             =   661
            Appearance      =   0
            AutoTickFrequency=   1
            AutoTickMarks   =   -1  'True
            BackColor       =   -2147483633
            BackgroundDrawMode=   0
            BorderStyle     =   0
            CurrentPosition =   50
            DisabledEvents  =   779
            DontRedraw      =   0   'False
            DownIsLeft      =   -1  'True
            Enabled         =   -1  'True
            HoverTime       =   -1
            LargeStepWidth  =   -1
            Maximum         =   400
            Minimum         =   1
            MousePointer    =   0
            Orientation     =   0
            ProcessContextMenuKeys=   -1  'True
            RangeSelectionEnd=   0
            RangeSelectionStart=   0
            RegisterForOLEDragDrop=   0   'False
            Reversed        =   0   'False
            RightToLeftLayout=   0   'False
            SelectionType   =   0
            ShowSlider      =   -1  'True
            SliderLength    =   -1
            SmallStepWidth  =   1
            SupportOLEDragImages=   -1  'True
            TickMarksPosition=   0
            ToolTipPosition =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Intensity:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Tag             =   "tab1"
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Tag             =   "tab1"
            Top             =   1080
            Width           =   585
         End
      End
      Begin BtnCtlsLibUCtl.Frame frmTab2 
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2535
         _cx             =   4471
         _cy             =   2778
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         BorderVisible   =   -1  'True
         ContentType     =   0
         DisabledEvents  =   3
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   0
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         MousePointer    =   0
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         IconIndexes     =   "frmMain.frx":3CF5
         Text            =   "frmMain.frx":3D2D
         Begin BtnCtlsLibUCtl.CommandButton cmdStopRender 
            Cancel          =   -1  'True
            Height          =   255
            Left            =   2140
            TabIndex        =   37
            ToolTipText     =   "Stop rendering"
            Top             =   480
            Width           =   255
            _cx             =   450
            _cy             =   450
            Appearance      =   0
            BackColor       =   -2147483633
            BorderStyle     =   0
            ButtonType      =   0
            ContentType     =   0
            DisabledEvents  =   1289
            DontRedraw      =   0   'False
            DropDownArrowHeight=   0
            DropDownArrowWidth=   15
            DropDownGlyph   =   54
            DropDownOnRight =   -1  'True
            DropDownPushed  =   0   'False
            DropDownStyle   =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            HAlignment      =   1
            HoverTime       =   -1
            IconAlignment   =   0
            IconMarginBottom=   0
            IconMarginLeft  =   0
            IconMarginRight =   0
            IconMarginTop   =   0
            KeepDropDownArrowAspectRatio=   -1  'True
            MousePointer    =   0
            MultiLine       =   -1  'True
            ProcessContextMenuKeys=   -1  'True
            Pushed          =   0   'False
            RegisterForOLEDragDrop=   0   'False
            RightToLeft     =   0
            ShowRightsElevationIcon=   0   'False
            ShowSplitLine   =   -1  'True
            Style           =   0
            SupportOLEDragImages=   -1  'True
            TextMarginBottom=   1
            TextMarginLeft  =   1
            TextMarginRight =   1
            TextMarginTop   =   1
            UseImprovedImageListSupport=   0   'False
            UseSystemFont   =   0   'False
            VAlignment      =   1
            IconIndexes     =   "frmMain.frx":3D4D
            CommandLinkNote =   "frmMain.frx":3D85
            Text            =   "frmMain.frx":3DA5
         End
         Begin ProgBarLibUCtl.ProgressBar ProgressBar 
            Height          =   255
            Left            =   120
            Top             =   480
            Width           =   1935
            _cx             =   3413
            _cy             =   450
            ActivateMarquee =   0   'False
            Appearance      =   3
            BackColor       =   -1
            BarColor        =   -1
            BarStyle        =   0
            BorderStyle     =   0
            CurrentValue    =   0
            DisabledEvents  =   3
            DontRedraw      =   0   'False
            Enabled         =   -1  'True
            HoverTime       =   -1
            MarqueeStepDuration=   50
            Maximum         =   100
            Minimum         =   0
            MousePointer    =   0
            Orientation     =   0
            ProgressState   =   1
            RegisterForOLEDragDrop=   0   'False
            RightToLeftLayout=   0   'False
            SmoothReverse   =   -1  'True
            StepWidth       =   10
            SupportOLEDragImages=   -1  'True
         End
         Begin BtnCtlsLibUCtl.CommandButton cmdSave 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1575
            _cx             =   2778
            _cy             =   661
            Appearance      =   0
            BackColor       =   -2147483633
            BorderStyle     =   0
            ButtonType      =   0
            ContentType     =   0
            DisabledEvents  =   1289
            DontRedraw      =   0   'False
            DropDownArrowHeight=   0
            DropDownArrowWidth=   15
            DropDownGlyph   =   54
            DropDownOnRight =   -1  'True
            DropDownPushed  =   0   'False
            DropDownStyle   =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            HAlignment      =   1
            HoverTime       =   -1
            IconAlignment   =   0
            IconMarginBottom=   0
            IconMarginLeft  =   0
            IconMarginRight =   0
            IconMarginTop   =   0
            KeepDropDownArrowAspectRatio=   -1  'True
            MousePointer    =   0
            MultiLine       =   -1  'True
            ProcessContextMenuKeys=   -1  'True
            Pushed          =   0   'False
            RegisterForOLEDragDrop=   0   'False
            RightToLeft     =   0
            ShowRightsElevationIcon=   0   'False
            ShowSplitLine   =   -1  'True
            Style           =   0
            SupportOLEDragImages=   -1  'True
            TextMarginBottom=   1
            TextMarginLeft  =   1
            TextMarginRight =   1
            TextMarginTop   =   1
            UseImprovedImageListSupport=   0   'False
            UseSystemFont   =   -1  'True
            VAlignment      =   1
            IconIndexes     =   "frmMain.frx":3DC7
            CommandLinkNote =   "frmMain.frx":3DFF
            Text            =   "frmMain.frx":3E1F
         End
         Begin VB.Label lblProgress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nothing rendered yet..."
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1620
         End
      End
      Begin BtnCtlsLibUCtl.Frame frmTab3 
         Height          =   5295
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2535
         _cx             =   4471
         _cy             =   9340
         Appearance      =   0
         BackColor       =   -2147483633
         BorderStyle     =   0
         BorderVisible   =   -1  'True
         ContentType     =   0
         DisabledEvents  =   3
         DontRedraw      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         HAlignment      =   0
         HoverTime       =   -1
         IconAlignment   =   0
         IconMarginBottom=   0
         IconMarginLeft  =   0
         IconMarginRight =   0
         IconMarginTop   =   0
         MousePointer    =   0
         RegisterForOLEDragDrop=   0   'False
         RightToLeft     =   0
         Style           =   0
         SupportOLEDragImages=   -1  'True
         UseImprovedImageListSupport=   0   'False
         UseSystemFont   =   -1  'True
         IconIndexes     =   "frmMain.frx":3E67
         Text            =   "frmMain.frx":3E9F
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Freeware"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   4320
            Width           =   2295
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hot Spot Studio v2.1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   13
            Top             =   390
            Width           =   1455
         End
         Begin VB.Image imgRLV 
            Height          =   2250
            Left            =   405
            MouseIcon       =   "frmMain.frx":3EBF
            Picture         =   "frmMain.frx":41C9
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblWeb 
            BackStyle       =   0  'Transparent
            Caption         =   "http://www.rlvision.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   405
            TabIndex        =   12
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "developed by"
            Height          =   225
            Left            =   120
            TabIndex        =   11
            Top             =   1065
            Width           =   2295
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Visit my homepage for more cool software!"
            Height          =   375
            Left            =   480
            TabIndex        =   10
            Top             =   4680
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright (c) 2002-2010"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   3960
            Width           =   2295
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   360
            Picture         =   "frmMain.frx":10943
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin TabStripCtlLibUCtl.TabStrip TabStrip2 
      Height          =   5295
      Left            =   3000
      TabIndex        =   24
      Top             =   390
      Width           =   5055
      _cx             =   8916
      _cy             =   9340
      AllowDragDrop   =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      CloseableTabs   =   0   'False
      DisabledEvents  =   262379
      DragActivateTime=   -1
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      FixedTabWidth   =   96
      FocusOnButtonDown=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HorizontalPadding=   6
      HotTracking     =   0   'False
      HoverTime       =   -1
      InsertMarkColor =   0
      MinTabWidth     =   -1
      MousePointer    =   0
      MultiRow        =   0   'False
      MultiSelect     =   0   'False
      OLEDragImageStyle=   0
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      RaggedTabRows   =   -1  'True
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ScrollToOpposite=   0   'False
      ShowButtonSeparators=   -1  'True
      ShowToolTips    =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TabBoundingBoxDefinition=   8198
      TabCaptionAlignment=   0
      TabHeight       =   18
      TabPlacement    =   0
      UseFixedTabWidth=   0   'False
      UseSystemFont   =   -1  'True
      VerticalPadding =   3
   End
   Begin BtnCtlsLibUCtl.CommandButton cmdBackToDesign 
      Height          =   495
      Left            =   240
      TabIndex        =   36
      ToolTipText     =   "Make a high quality rendering of this scene"
      Top             =   7680
      Width           =   1335
      Visible         =   0   'False
      _cx             =   2355
      _cy             =   873
      Appearance      =   0
      BackColor       =   -2147483633
      BorderStyle     =   0
      ButtonType      =   0
      ContentType     =   0
      DisabledEvents  =   1289
      DontRedraw      =   0   'False
      DropDownArrowHeight=   0
      DropDownArrowWidth=   15
      DropDownGlyph   =   54
      DropDownOnRight =   -1  'True
      DropDownPushed  =   0   'False
      DropDownStyle   =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   1
      HoverTime       =   -1
      IconAlignment   =   0
      IconMarginBottom=   0
      IconMarginLeft  =   0
      IconMarginRight =   0
      IconMarginTop   =   0
      KeepDropDownArrowAspectRatio=   -1  'True
      MousePointer    =   0
      MultiLine       =   -1  'True
      ProcessContextMenuKeys=   -1  'True
      Pushed          =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ShowRightsElevationIcon=   0   'False
      ShowSplitLine   =   -1  'True
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextMarginBottom=   1
      TextMarginLeft  =   1
      TextMarginRight =   1
      TextMarginTop   =   1
      UseImprovedImageListSupport=   0   'False
      UseSystemFont   =   -1  'True
      VAlignment      =   1
      IconIndexes     =   "frmMain.frx":10EC5
      CommandLinkNote =   "frmMain.frx":10EFD
      Text            =   "frmMain.frx":10F1D
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.rlvision.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   6720
      TabIndex        =   35
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private currX As Integer
Private currY As Integer
Private currRGB As Long
Private myLights() As Spotlight
Private myBackup() As Spotlight
Private check1_LastValue As Integer

Private bJustStarted As Boolean

Private Type Spotlight
    X As Integer
    relX As Double
    Y As Integer
    relY As Double
    r As Integer
    G As Integer
    b As Integer
    i As Integer    'intensity
    s As Integer    'scale
    Type As Integer '0=normal, 1=negative, 2=negative absolute
End Type

Private iHovered As Integer
Private iSelected As Integer
Private iMoving As Integer
Private bMoved As Boolean
Private iNewLight As Integer
Private bSelectingNewLight As Boolean
Private iLastWidth As Integer
Private iLastHeight As Integer
Private bStopFlag As Boolean
Public BgColorRGB As Long
Private lastSaveName As String
Private iFilterIndex As Long

Private lastMouseDownX As Single
Private lastMouseDownY As Single
Private lastMouseDownButton As Integer
Private lastMouseDownShift As Integer

Private hImgLst_StopRenderButton As Long

Private Sub Form_Load()
    
    ReDim myLights(0)
    
    currRGB = 4419583
    picCurrColor.BackColor = currRGB
    
    backRGB = RGB(0, 0, 0)
    Picture1.BackColor = backRGB
    picRender.BackColor = backRGB

    iMoving = -1
    iSelected = -1
    iHovered = -1
    bMoved = False
    iNewLight = -1

    cmdSave.Enabled = False
    Command7.Enabled = False
    cmdRender.Enabled = False
    cmdStopRender.Enabled = False

    TabStrip1.Tabs.Add ("Design")
    TabStrip1.Tabs.Add ("Rendering")
    TabStrip1.Tabs.Add ("About")

    SetIcon Me.hWnd, "AAA", True

    frmMain.Width = 8745
    frmMain.Height = 6650

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnloadXpApp
End Sub


Private Sub Check1_SelectionStateChanged(ByVal previousSelectionState As BtnCtlsLibUCtl.SelectionStateConstants, ByVal newSelectionState As BtnCtlsLibUCtl.SelectionStateConstants)
    Call UpdateLight(iSelected)
End Sub

Private Sub cmdBackToDesign_Click(ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)
    Set TabStrip1.ActiveTab = TabStrip1.Tabs(0)
End Sub

Private Sub cmdLoadBG_Click(ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim d As New cCommonDialog
    Dim fil As String
    Call d.VBGetOpenFileName(fil, , , , , , "Pictures (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif;")

    If fil <> "" Then
        Picture1.Picture = LoadPicture(fil)
        picRender.Picture = LoadPicture(fil)
    End If

End Sub

Private Sub cmdSave_Click(ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim d As New cCommonDialog
    Dim fil As String
    fil = lastSaveName
    If iFilterIndex = 0 Then iFilterIndex = 1
    Call d.VBGetSaveFileName(fil, , , "JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|BMP (*.bmp)|*.bmp|", iFilterIndex)

    If fil <> "" Then
        ret = False
        If iFilterIndex = 1 Then
            If LCase(Right(fil, 4)) <> ".jpg" And LCase(Right(fil, 5)) <> ".jpeg" Then fil = fil & ".jpg"
            ret = SavePictureEx(picRender.Image, fil, FIF_JPEG, FISO_JPEG_QUALITYSUPERB)
        ElseIf iFilterIndex = 2 Then
            If LCase(Right(fil, 4)) <> ".png" Then fil = fil & ".png"
            ret = SavePictureEx(picRender.Image, fil, FIF_PNG, , FICD_24BPP)
        ElseIf iFilterIndex = 3 Then
            If LCase(Right(fil, 4)) <> ".tif" And LCase(Right(fil, 5)) <> ".tiff" Then fil = fil & ".tif"
            ret = SavePictureEx(picRender.Image, fil, FIF_TIFF, , FICD_24BPP)
        ElseIf iFilterIndex = 4 Then
            If LCase(Right(fil, 4)) <> ".bmp" Then fil = fil & ".bmp"
            ret = SavePictureEx(picRender.Image, fil, FIF_BMP, , FICD_24BPP)
        End If
        lastSaveName = fil
    End If

End Sub

Private Sub cmdSetBG_Click(ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)

    frmSetBgCol.Show vbModal

    Picture1.BackColor = BgColorRGB
    picRender.BackColor = BgColorRGB

End Sub

Private Sub cmdStopRender_Click(ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)
    bStopFlag = True
End Sub

Private Sub Command3_Click()

    For n = 1 To UBound(myLights)
        Unload shpLight(n)
    Next
    
    ReDim myLights(0)
    Picture1.Cls
    Picture1.Picture = Nothing
    picRender.Cls
    picRender.Picture = Nothing
    
    cmdRender.Enabled = False

End Sub

Private Sub cmdRender_Click(ByVal button As Integer, ByVal shift As Integer, ByVal cx As Single, ByVal cy As Single)

    If UBound(myLights) = 0 Then Exit Sub

    Set TabStrip1.ActiveTab = TabStrip1.Tabs(1)
    
    doMax = 100
    doCount = 0

    T = Timer

    picRender.Cls
    ProgressBar.CurrentValue = 0

    bStopFlag = False
    cmdStopRender.Enabled = True

    Screen.MousePointer = 13
    
    cmdSave.Enabled = False
    TrackBar_Size.Enabled = False
    TrackBar_Int.Enabled = False
    txtSize.Enabled = False
    txtInt.Enabled = False
    cmdRender.Enabled = False
    Check1.Enabled = False
    Picture2.Enabled = False
    cmdSetBG.Enabled = False
    cmdLoadBG.Enabled = False
    
    Dim X As Long, Y As Long
    Dim n As Long, xx As Long, yy As Long
    Dim origR As Double, origG As Double, origB As Double, newR As Double, newG As Double, newB As Double
    Dim r As Double, G As Double, b As Double, d As Double, v As Double

    ' theory:
    '   v = exp( - ( d * d ))
    '   d is the distance from middle
    '   v becomes between 0 and 1
    '   scale d to experiment with size
    '


    For Y = 0 To picRender.Height
        For X = 0 To picRender.Width
            
            origR = (picRender.Point(X, Y) And 255)
            origG = (picRender.Point(X, Y) And 65280) / 256
            origB = (picRender.Point(X, Y) And 16711680) / 256 / 256

            newR = 0
            newG = 0
            newB = 0
            
            For n = 1 To UBound(myLights)

                'get distance to light
                xx = X - myLights(n).X
                yy = Y - myLights(n).Y
                d = Sqr((xx * xx) + (yy * yy))
    
                d = d / myLights(n).s  'scale distance
    
                v = Exp(-(d * d))   'gaussian style. v => between 0 and 1
                
                r = v * myLights(n).r
                r = r * myLights(n).i
                G = v * myLights(n).G
                G = G * myLights(n).i
                b = v * myLights(n).b
                b = b * myLights(n).i

                Select Case myLights(n).Type
                Case 0
                    newR = newR + r
                    newG = newG + G
                    newB = newB + b
                Case 1
                    newR = newR - r
                    newG = newG - G
                    newB = newB - b
                End Select

            Next

            newR = Abs(newR)
            If newR > 255 Then newR = 255
            newG = Abs(newG)
            If newG > 255 Then newG = 255
            newB = Abs(newB)
            If newB > 255 Then newB = 255
            
            picRender.PSet (X, Y), RGB(origR + newR, origG + newG, origB + newB)
            
            doCount = doCount + 1
            If doCount > doMax Then
                doCount = 0
                If bStopFlag = True Then
                    lblProgress = "Stopped..."
                    GoTo stopp
                End If
                DoEvents
            End If
        
        Next
        
        lblProgress = "Rendering picture... " & Int((Y / picRender.Height) * 100) & "%"
        
        ProgressBar.CurrentValue = Int((Y / picRender.Height) * 100)

    Next

    lblProgress = "Finished!"
    cmdSave.Enabled = True

stopp:

    Screen.MousePointer = 0

    TrackBar_Size.Enabled = True
    TrackBar_Int.Enabled = True
    txtSize.Enabled = True
    txtInt.Enabled = True
    cmdRender.Enabled = True
    Check1.Enabled = True
    Picture2.Enabled = True
    cmdSetBG.Enabled = True
    cmdLoadBG.Enabled = True
    
    cmdStopRender.Enabled = False

    Debug.Print "Render time: " & Round(Timer - T, 1) & "s"

End Sub

Private Sub Command6_Click()

    ReDim myBackup(UBound(myLights))

    For n = 1 To UBound(myLights)
        myBackup(n).b = myLights(n).b
        myBackup(n).G = myLights(n).G
        myBackup(n).i = myLights(n).i
        myBackup(n).r = myLights(n).r
        myBackup(n).s = myLights(n).s
        myBackup(n).Type = myLights(n).Type
        myBackup(n).X = myLights(n).X
        myBackup(n).Y = myLights(n).Y
    Next

    Command7.Enabled = True

End Sub

Private Sub Command7_Click()

    ReDim myLights(UBound(myBackup))

    For n = 1 To UBound(myBackup)
        myLights(n).b = myBackup(n).b
        myLights(n).G = myBackup(n).G
        myLights(n).i = myBackup(n).i
        myLights(n).r = myBackup(n).r
        myLights(n).s = myBackup(n).s
        myLights(n).Type = myBackup(n).Type
        myLights(n).X = myBackup(n).X
        myLights(n).Y = myBackup(n).Y
    Next

End Sub

Private Sub Command8_Click()

    Picture1.BackColor = picCurrColor.BackColor
    picRender.BackColor = picCurrColor.BackColor

End Sub

Private Sub Form_Resize()

    If cmdSetBG.Enabled = False Then Exit Sub

    If frmMain.WindowState = 1 Then Exit Sub 'minimized

    picRender.Cls
    If cmdSave.Enabled = True And TabStrip1.ActiveTab.Index = 1 Then
        Set TabStrip1.ActiveTab = TabStrip1.Tabs(0)
        cmdSave.Enabled = False
        lblProgress = "Please render again..."
    End If
    
    Picture1.Left = TabStrip2.Left + 4
    Picture1.Top = TabStrip2.Top + 6
    picRender.Left = Picture1.Left
    picRender.Top = Picture1.Top

    If frmMain.Width < 8745 Then frmMain.Width = 8745
    If frmMain.Height < 6615 Then frmMain.Height = 6615
        

    TabStrip2.Width = frmMain.ScaleWidth - TabStrip2.Left - 8
    TabStrip2.Height = frmMain.ScaleHeight - TabStrip2.Top - 8

    T = 0
    If IsThemed() Then T = 2

    If frmMain.Width >= 8745 Then
        Picture1.Width = TabStrip2.Width - 8 - T
        picRender.Width = Picture1.Width
    
        For n = 1 To UBound(myLights)
            myLights(n).X = Picture1.Width * myLights(n).relX
            shpLight(n).Left = myLights(n).X - (shpLight(n).Width / 2)
        Next
    End If
    
    If frmMain.Height >= 6615 Then
        Picture1.Height = TabStrip2.Height - 10 - T
        picRender.Height = Picture1.Height
    
    
        For n = 1 To UBound(myLights)
            myLights(n).Y = Picture1.Height * myLights(n).relY
            shpLight(n).Top = myLights(n).Y - (shpLight(n).Height / 2)
        Next
    End If


    lblWeb(1).Left = frmMain.ScaleWidth - lblWeb(1).Width - 10

End Sub

Private Sub HScroll1_Change()
    txtSize = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    txtSize = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
    txtInt = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
    txtInt = HScroll2.Value
End Sub

Private Sub imgRLV_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.rlvision.com/script/redirect.asp?app=spots", vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub imgRLV_MouseMove(button As Integer, shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)   'set hand cursor
End Sub

Private Sub lblWeb_Click(Index As Integer)
    Call ShellExecute(Me.hWnd, "open", "http://www.rlvision.com/script/redirect.asp?app=spots", vbNullString, vbNullString, SW_NORMAL)
End Sub
Private Sub lblWeb_MouseMove(Index As Integer, button As Integer, shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, IDC_HAND)   'set hand cursor

    lblWeb(Index).ForeColor = RGB(255, 0, 0)
    tmrRestoreLinkLabels.Enabled = True

End Sub

Private Sub picRender_Click()
    Set TabStrip1.ActiveTab = TabStrip1.Tabs(0)
End Sub

Private Sub Picture2_MouseMove(button As Integer, shift As Integer, X As Single, Y As Single)

    If button = 1 Then

        If X > (1 * Screen.TwipsPerPixelX) And Y > (1 * Screen.TwipsPerPixelY) And X < Picture2.Width - (1 * Screen.TwipsPerPixelX) And Y < Picture2.Height - (1 * Screen.TwipsPerPixelY) Then
            currRGB = Picture2.Point(X, Y)
            If currRGB >= 0 Then
                picCurrColor.BackColor = currRGB
                Debug.Print picCurrColor.BackColor
                Call UpdateLight(iSelected)
            End If
        End If
    
    End If

End Sub

Private Sub Picture2_MouseUp(button As Integer, shift As Integer, X As Single, Y As Single)

    If button = 1 Then

        If X > 0 And Y > 0 And X < Picture2.Width And Y < Picture2.Height Then
            currRGB = Picture2.Point(X, Y)
            picCurrColor.BackColor = currRGB
            Call UpdateLight(iSelected)
        End If
    
    End If

End Sub


Private Sub TabStrip1_ActiveTabChanged(ByVal previousActiveTab As TabStripCtlLibUCtl.ITabStripTab, ByVal newActiveTab As TabStripCtlLibUCtl.ITabStripTab)

    i = newActiveTab.Index

    frmTab11.Visible = False
    frmTab12.Visible = False
    frmTab2.Visible = False
    frmTab3.Visible = False
    cmdRender.Visible = False
    frmHelp.Visible = False
    
    If i = 0 Then
        Picture1.ZOrder
        frmTab11.Visible = True
        frmTab12.Visible = True
        frmTab11.ZOrder
        frmTab12.ZOrder
        cmdRender.Visible = True
        cmdRender.ZOrder
        frmHelp.Visible = True
        frmHelp.ZOrder
    ElseIf i = 1 Then
        picRender.ZOrder
        frmTab2.Visible = True
        frmTab2.ZOrder
    ElseIf i = 2 Then
        frmTab3.Visible = True
        frmTab3.ZOrder
    End If

End Sub

Private Sub TabStrip1_ActiveTabChanging(ByVal previousActiveTab As TabStripCtlLibUCtl.ITabStripTab, cancelChange As Boolean)
    If cmdSetBG.Enabled = False Then cancelChange = True
End Sub

Private Sub TrackBar_Int_PositionChanged(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
    txtInt = TrackBar_Int.CurrentPosition
    Call UpdateLight(iSelected)
End Sub

Private Sub TrackBar_Size_PositionChanged(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
    txtSize = TrackBar_Size.CurrentPosition
    Call UpdateLight(iSelected)
End Sub

Private Sub txtSize_LostFocus()
    txtSize = Val(txtSize)
    Call UpdateLight(iSelected)
End Sub
Private Sub txtInt_LostFocus()
    txtInt = Val(txtInt)
    Call UpdateLight(iSelected)
End Sub

Private Sub Picture1_MouseUp(button As Integer, shift As Integer, X As Single, Y As Single)

    iMoving = -1

End Sub


Private Sub Picture1_MouseMove(button As Integer, shift As Integer, X As Single, Y As Single)

    If cmdSetBG.Enabled = False Then Exit Sub

    If iNewLight <> -1 And iSelected = -1 Then
        If button = 1 Then
            iMoving = iNewLight
            iSelected = iNewLight
        Else
            iNewLight = -1
        End If
    End If

    If iMoving <> -1 Then
        bMoved = True
        myLights(iSelected).X = X
        myLights(iSelected).Y = Y
        shpLight(iSelected).Left = myLights(iSelected).X - (shpLight(iSelected).Width / 2)
        shpLight(iSelected).Top = myLights(iSelected).Y - (shpLight(iSelected).Height / 2)
        myLights(iSelected).relX = X / Picture1.Width
        myLights(iSelected).relY = Y / Picture1.Height
        
        Exit Sub
    End If

    iHovered = -1
    For n = 1 To UBound(myLights)
        rad = shpLight(n).Width / 2
        If rad < 16 Then rad = 16
        'rad = 10
        
        dx = X - myLights(n).X
        dy = Y - myLights(n).Y
        If Sqr((dx * dx) + (dy * dy)) < rad Then
            If n <> iSelected Then shpLight(n).BorderColor = vbWhite
            iHovered = n
        Else
            If n <> iSelected Then shpLight(n).FillStyle = 1
            If n <> iSelected Then shpLight(n).BorderColor = shpLight(n).FillColor
        End If
    Next

End Sub

Private Sub Picture1_DblClick()
    'forward click since dblclick consumes next mouse down
    Call Picture1_MouseDown(lastMouseDownButton, lastMouseDownShift, lastMouseDownX, lastMouseDownY)
End Sub

Private Sub Picture1_MouseDown(button As Integer, shift As Integer, X As Single, Y As Single)

lastMouseDownX = X
lastMouseDownY = Y
lastMouseDownButton = button
lastMouseDownShift = shift

Set TabStrip1.ActiveTab = TabStrip1.Tabs(0)

    If cmdSetBG.Enabled = False Then Exit Sub

    currX = X
    currY = Y

    Dim n As Integer
    
    If iMoving <> -1 Then
        Exit Sub
    End If
    
    iMoving = -1
    If button = 1 And iHovered <> -1 Then
        'select/move
        iSelected = iHovered
        iMoving = iSelected
        bMoved = False
    
        For n = 1 To UBound(myLights)
            If n <> iSelected Then shpLight(n).FillStyle = 1
        Next
    
        shpLight(iSelected).FillStyle = 0
        shpLight(iSelected).BorderColor = shpLight(iSelected).FillColor
    
    
        bSelectingNewLight = True
        picCurrColor.BackColor = RGB(myLights(iSelected).r, myLights(iSelected).G, myLights(iSelected).b)
        currRGB = picCurrColor.BackColor
        txtSize = myLights(iSelected).s
        TrackBar_Size.CurrentPosition = txtSize
        txtInt = myLights(iSelected).i
        TrackBar_Int.CurrentPosition = txtInt
        Check1.SelectionState = myLights(iSelected).Type
        bSelectingNewLight = False
        shpLight(iSelected).ZOrder

    ElseIf button = 1 And iHovered = -1 Then
        'unselect
        
        iSelected = -1
        iHovered = -1
        
        If UBound(myLights) = 0 Then
            Call MsgBox("Click the RIGHT mouse button to add a light to the canvas!", vbInformation)
        End If
        
    End If

    If button = 2 And iHovered = -1 Then

        'new light
        If lblStartupTip.Visible Then lblStartupTip.Visible = False

        ReDim Preserve myLights(UBound(myLights) + 1)
    
        n = UBound(myLights)
    
        myLights(n).X = currX
        myLights(n).X = currX
        myLights(n).relX = currX / Picture1.Width
        myLights(n).Y = currY
        myLights(n).relY = currY / Picture1.Height

        Load shpLight(n)
        shpLight(n).Visible = True
        shpLight(n).ZOrder
        
        cmdRender.Enabled = True
        
        iSelected = n
        iHovered = n
        iNewLight = n
        
        shpLight(iSelected).FillStyle = 0
        
        Call UpdateLight(n)
        
    ElseIf button = 2 And iHovered <> -1 Then

        'delete
        n = iHovered

        If n <> -1 And n <= UBound(myLights) Then
            myLights(n).X = myLights(UBound(myLights)).X
            myLights(n).Y = myLights(UBound(myLights)).Y
            myLights(n).i = myLights(UBound(myLights)).i
            myLights(n).s = myLights(UBound(myLights)).s
            myLights(n).r = myLights(UBound(myLights)).r
            myLights(n).G = myLights(UBound(myLights)).G
            myLights(n).b = myLights(UBound(myLights)).b
            myLights(n).Type = myLights(UBound(myLights)).Type


            shpLight(n).BorderColor = shpLight(UBound(myLights)).BorderColor
            shpLight(n).FillColor = shpLight(UBound(myLights)).FillColor
            shpLight(n).Width = shpLight(UBound(myLights)).Width
            shpLight(n).Height = shpLight(UBound(myLights)).Height
            shpLight(n).Left = shpLight(UBound(myLights)).Left
            shpLight(n).Top = shpLight(UBound(myLights)).Top

            Unload shpLight(UBound(myLights))
            ReDim Preserve myLights(UBound(myLights) - 1)

        End If

        If UBound(myLights) = 0 Then cmdRender.Enabled = False
    
    End If

End Sub


Private Sub UpdateLight(n)

    If n > -1 And n <= UBound(myLights) And bSelectingNewLight = False Then
    
        myLights(n).i = Val(txtInt)
        myLights(n).s = Val(txtSize)
        myLights(n).r = currRGB And 255
        myLights(n).G = (currRGB And 65280) / 256
        myLights(n).b = (currRGB And 16711680) / 256 / 256
        myLights(n).Type = Check1.SelectionState
    
        shpLight(n).BorderColor = RGB(myLights(n).r, myLights(n).G, myLights(n).b)
        shpLight(n).FillColor = shpLight(n).BorderColor
        If myLights(n).s > 10 Then
            shpLight(n).Width = myLights(n).s
            shpLight(n).Height = myLights(n).s
        Else
            shpLight(n).Width = 10
            shpLight(n).Height = 10
        End If
        shpLight(n).Left = myLights(n).X - (shpLight(n).Width / 2)
        shpLight(n).Top = myLights(n).Y - (shpLight(n).Height / 2)
    
    End If

End Sub


Private Sub tmrRestoreLinkLabels_Timer()

    Dim nActive As Integer: Dim i As Integer: Dim myX As Integer: Dim myY As Integer

    nActive = lblWeb.Count
    For i = lblWeb.LBound To lblWeb.UBound
        If lblWeb(i).ForeColor <> 12582912 Then
            myX = MouseX(lblWeb(i).Container.hWnd)
            myY = MouseY(lblWeb(i).Container.hWnd)
            
            If i = 0 Then
                myX = myX * Screen.TwipsPerPixelX
                myY = myY * Screen.TwipsPerPixelY
            End If

            If myY < lblWeb(i).Top _
            Or myY > lblWeb(i).Top + lblWeb(i).Height _
            Or myX < lblWeb(i).Left _
            Or myX > lblWeb(i).Left + lblWeb(i).Width _
            Then
                    lblWeb(i).ForeColor = 12582912
                    nActive = nActive - 1
            End If
        Else
            nActive = nActive - 1
        End If
    Next
    If nActive = 0 Then tmrRestoreLinkLabels.Enabled = False

End Sub

