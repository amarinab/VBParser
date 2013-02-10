VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "AFORATE"
   ClientHeight    =   9675
   ClientLeft      =   3615
   ClientTop       =   1575
   ClientWidth     =   12900
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox fechaAux 
      DataField       =   "FECHAMED"
      DataSource      =   "AFOROS"
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox c4Text 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox c3Text 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox c2Text 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox c1Text 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox traText 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3600
      Width           =   6255
   End
   Begin VB.TextBox delText 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox tipoText 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox pkText 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox locText 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2880
      Width           =   3855
   End
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   300
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton BajaEst 
      Caption         =   "Baja Estación"
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
      Left            =   3480
      Picture         =   "main.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Da de baja una estación"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton AltaEst 
      Caption         =   "Alta Estación"
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
      Left            =   2400
      Picture         =   "main.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Da de alta una nueva estación"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      Picture         =   "main.frx":205E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Muestra información de la Base de Datos"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton abrirBD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      Picture         =   "main.frx":2928
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Abre o Guarda la Base de Datos"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton imprimir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      Picture         =   "main.frx":37F2
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Imprime los datos de nuestra Base de Datos"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton restablecer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Picture         =   "main.frx":40BC
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Restablece los datos depués de un filtrado"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox estText 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   2520
      Width           =   6255
   End
   Begin VB.CommandButton buscar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      Picture         =   "main.frx":4986
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Filtra los datos que se mostrarán por la estación indicada"
      Top             =   240
      Width           =   975
   End
   Begin MSACAL.Calendar calendar 
      Bindings        =   "main.frx":5250
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   2775
      Left            =   8520
      TabIndex        =   32
      ToolTipText     =   "Filtra los datos a partir de la fecha indicada (y a partir de la estación si se indicó previamente)"
      Top             =   2160
      Width           =   4095
      _Version        =   524288
      _ExtentX        =   7223
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   14933984
      Year            =   1900
      Month           =   1
      Day             =   1
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton salir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12000
      Picture         =   "main.frx":526C
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Salir de la aplicación"
      Top             =   240
      Width           =   600
   End
   Begin MSDataGridLib.DataGrid vista 
      Bindings        =   "main.frx":6136
      Height          =   3855
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "IDEST"
         Caption         =   "Estación"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FECHAMED"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "HORAMED"
         Caption         =   "Hora"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "TOTVEH1"
         Caption         =   "TotVeh"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "VELMED1"
         Caption         =   "VelMed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "VEHPES1"
         Caption         =   "VehPes"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "TOTVEH2"
         Caption         =   "TotVeh"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "VELMED2"
         Caption         =   "VelMed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "VEHPES2"
         Caption         =   "VehPes"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "TOTVEHT"
         Caption         =   "TotVeh"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "VELMEDT"
         Caption         =   "VelMed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "VEHPEST"
         Caption         =   "VehPes"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column04 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column07 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column08 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column09 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column10 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column11 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   900,284
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AFOROS 
      Height          =   375
      Left            =   240
      ToolTipText     =   "Navega por los registros"
      Top             =   5040
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   1
      CommandTimeout  =   1
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AFOROS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton importar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Picture         =   "main.frx":614B
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Importa el fichero de texto seleccionado a nuestra base de datos"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox ruta 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   1200
      Width           =   11295
   End
   Begin VB.CommandButton abrir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Picture         =   "main.frx":6A15
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Abre un fichero para importarlo"
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      DragIcon        =   "main.frx":72DF
      Height          =   270
      Left            =   0
      TabIndex        =   25
      Top             =   9405
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17092
            Text            =   "AFORATE version 1.0"
            TextSave        =   "AFORATE version 1.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/05/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "14:03"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Fichero"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9345
      TabIndex        =   37
      Top             =   5160
      Width           =   2700
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Canal 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   36
      Top             =   5160
      Width           =   2700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Canal 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3945
      TabIndex        =   35
      Top             =   5160
      Width           =   2700
   End
   Begin VB.Label c4 
      Alignment       =   1  'Right Justify
      Caption         =   "Canal 4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label c3 
      Alignment       =   1  'Right Justify
      Caption         =   "Canal 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label c2 
      Alignment       =   1  'Right Justify
      Caption         =   "Canal 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label c1 
      Alignment       =   1  'Right Justify
      Caption         =   "Canal 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label tra 
      Alignment       =   1  'Right Justify
      Caption         =   "Tramo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label del 
      Alignment       =   1  'Right Justify
      Caption         =   "Delegación"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label tipol 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label pk 
      Alignment       =   1  'Right Justify
      Caption         =   "PK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label loc 
      Alignment       =   1  'Right Justify
      Caption         =   "Localización"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      Caption         =   "Estación"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sPath As String
Public sName As String
Public SDir As String
Public todos As Integer



'############## DESPLEGA EL FORMULARIO DE ALTA DE ESTACION ##########
Private Sub AltaEst_Click()
    progreso.Value = 0
    Form1.Show
End Sub


'############## DESPLEGA EL FORMULARIO DE BAJA DE ESTACION ##########
Private Sub BajaEst_Click()
    progreso.Value = 0
    Form2.Show
End Sub


Private Sub fechaAux_Change()
    calendar.Value = fechaAux.Text
End Sub

'############## METODOS DE  INICIALIZACIÓN DEL FORMULARIO ##########


Private Sub Form_Initialize()
    AFOROS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseDatos.mdb;Persist Security Info=False"
    AFOROS.RecordSource = "ESTDATAUX"
End Sub



Private Sub Form_Load()
    AFOROS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseDatos.mdb;Persist Security Info=False"
    AFOROS.RecordSource = "ESTDATAUX"
    
    restab

    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 8700)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 9700)
    
salir:
    Exit Sub
err_Load:
    cnn.Close
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al cargar la aplicación")
    Resume salir
End Sub



'########### MÉTODO PARA SALIR DE LA APLICACION ###########3
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub




'######### ORDEN PARA ABRIR/GUARDAR LA BASE DE DATOS ##########
Private Sub abrirBD_Click()
    progreso.Value = 0
    Shell "Explorer.exe " & App.Path & "\BaseDatos.mdb", vbMaximizedFocus
End Sub


'########## MÉTODO PARA SELECCIONAR EL ARCHIVO A IMPORTAR ##########
Private Sub abrir_Click()
    abrirFichero
End Sub


Private Sub abrirFichero()
On Error GoTo ErrorProducido
    Dim sFile As String
    
    progreso.Value = 0
    
    With dlgCommonDialog
        .DialogTitle = "Abrir"
        .CancelError = False
        .Filter = "Archivos FIM|*.FIM"
        .InitDir = App.Path
        .ShowOpen
        
        sName = .FileTitle
        
        If Len(.fileName) = 0 Then
            Exit Sub
        End If
        sFile = .fileName
    End With
    
    ruta.Text = sFile
    sPath = Left(sFile, (Len(sFile) - Len(sName)) - 1)
    
    
salir:
    Exit Sub
ErrorProducido:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al abrir el archivo")
    Resume salir
End Sub


'####### METODO PARA IMPRIMIR (UN POCO EN EL AIRE) ###########
Private Sub imprimir_Click()
    On Error GoTo ErrorProducido
    progreso.Value = 0
    With CommonDialog1
        .DialogTitle = "Imprimir"
        .CancelError = False
        .fileName = App.Path & "/BaseDatos.mdb"
        .PrinterDefault = False
        .Filter = "Bases de Datos Access (*.mdb)|*.mdb"
        .ShowPrinter
    End With
salir:
    Exit Sub
ErrorProducido:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al imprimir")
    Resume salir

End Sub


'########### METODO IMPORTAR (ESENCIA DEL PROGRAMA) ##########
'En este método se da posibilidad de procesar todos los del directorio del fichero
'seleccionado. Se recorren pues todos los ficheros del directorio y se los manda al
'metodo siguiente "importarFichero" que es el que recorre cada fichero completamente
'importando los datos a la base de datos

Private Sub importar_Click()
    
    On Error GoTo Err_importar_Click
    
    If ruta.Text = "" Or ruta.Text = " " Then
        a = MsgBox("Antes, debe seleccionar un fichero", vbExclamation, "Error de Importación")
        progreso.Value = 0
        abrirFichero
        Exit Sub
    End If
    'Preguntamos si quiere importar todos los del directorio
    todos = MsgBox("¿Desea procesar todos los archivos *.FIM de este directorio?", vbOKCancel, "Procesar Todos")
    If todos = 2 Then
        importarArchivo (ruta.Text)
    Else
        'LISTAMOS LOS FICHEROS de la ruta
        Dim fso As Object 'New FileSystemObject
        Dim carpeta As Object 'Folder
        Dim archivos As Object 'Files
        Dim archivo As Object 'File
        Dim fileName, file, fileExt As String
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set carpeta = fso.GetFolder(sPath)
        Set archivos = carpeta.Files
    
        'importamos cada archivo (siempre y cuando sea *.fim)
        For Each archivo In archivos
            fileName = archivo.Name
            sName = fileName
            fileExt = Mid(fileName, Len(fileName) - 2, 3)
            If fileExt = "fim" Or fileExt = "FIM" Or fileExt = "Fim" Then
                file = sPath & "\" & fileName
                importarArchivo (file)
            End If
        Next
        
        Set archivos = Nothing
        Set carpeta = Nothing
        Set fso = Nothing
    End If
    ruta = ""
Exit_importar_Click:
        Exit Sub
Err_importar_Click:
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Importación")
        ruta = ""
        Resume Exit_importar_Click
End Sub


'####### METODO DE IMPORTACION ##########
'Recorre linea a linea, campo a campo el fichero que le pasamos como parametro
'y va importando los datos correctamente a nuestra base de datos
Private Sub importarArchivo(file As String)
    
    'Variables de base de datos
    Dim cnn As ADODB.Connection
    Dim regAux As ADODB.Recordset
    Dim baseC As Database
    Set baseC = OpenDatabase(App.Path & "\BaseDatos.mdb")
    
    'Variables a nivel de fichero
    Dim numBloques As Integer
    Dim numLineasBloque As Integer
    Dim totalLineas As Integer
    Dim linea As String
    Dim cont As Integer
    Dim QTencontrado As Boolean
    Dim QLencontrado As Boolean
    QTencontrado = False
    QLencontrado = False
    
    'Variables a nivel de linea
    Dim idEstacion As String
    Dim fecha As String
    'fecha = "12/28/06"
    Dim fechaAux As Date
    Dim horaInicio As Integer
    Dim hora As Integer
    Dim intervalo As Integer
    Dim canal As Integer
    Dim totVeh As Integer
    Dim velMed As Integer
    Dim totPes As Integer
    Dim limitePeso As String
    
    'Variables auxiliares
    Dim ini As Integer
    Dim aux As String
    Dim acumulado As Integer
    
    
    
    On Error GoTo Err_importarArchivo
              
        
        'abrimos el fichero
        Open file For Input Shared As #2
        
        'obtenemos el nombre de la estacion a partir del nombre del fichero
        Dim s As String
        For i = 1 To Len(sName)
            s = Mid(sName, i, 1)
            If s = "_" Then
                idEstacion = Mid(sName, 1, i - 1)
                Exit For
            End If
        Next
        
        
        
        ' Establezco la conexión con la base de datos actual
        Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            
            'Comprobamos que la estacion existe en nuestra base de datos
            consulta = "SELECT IDENTIFICACION FROM EST WHERE IDENTIFICACION='" & idEstacion & "';"
            Set regAux = .Execute(consulta)
            If regAux.EOF <> False Then
                a = MsgBox("La estación a la que hace referencia no está registrada" & vbCrLf & "¿Desea dar de alta a la estación '" & idEstacion & "' en este momento?", vbOKCancel, "Error de Importación")
                If a = 1 Then
                    Form1.Show
                    Form1.identificacionT.Text = idEstacion
                    cnn.Close
                    Close #2
                    Exit Sub
                Else
                    cnn.Close
                    Close #2
                    Exit Sub
                End If
            End If
            
            'Extraemos la primera linea para obtener informacion necesaria para recorrer el fichero
            Line Input #2, linea
            longitud = Len(linea)
            If (longitud <> 78) Then
                a = MsgBox("El fichero *.FIM no es correcto", vbOKCancel, "Error de Importación")
                Close #2
                Exit Sub
            End If
            
            cont = 1
            'MsgBox linea
            tipo = Mid(linea, 58, 2)
            'MsgBox tipo
            numBloques = Mid(linea, 76, 4)
            'MsgBox numBloques
            numLineasBloque = Mid(linea, 71, 4)
            'MsgBox numLineasBloque
            totalLineas = numBloques * numLineasBloque + numBloques
            'MsgBox totalLineas
            progreso.Max = totalLineas
            
            
            Do
                
                
                
                
                'Estos son los tipos de bloques que tenemos en cuenta
                If tipo = "QT" Or tipo = "QP" Or tipo = "VT" Or tipo = "QL" Or tipo = "LC" Then
                    
                    'Calculamos la clave primaria
                    fecha = Mid(linea, 38, 2) & "/" & Mid(linea, 33, 2) & "/" & Mid(linea, 28, 2)
                    'MsgBox fecha
                    horaInicio = Mid(linea, 43, 2)
                    'MsgBox horaInicio
                    canal = Mid(linea, 69, 1) + 1
                    'MsgBox canal
                    intervalo = Mid(linea, 51, 4)
                    'MsgBox intervalo
                
                
                    'Insertamos en la base de datos según el tipo del bloque
                    
                    
                    
                    'Conteo de vehiculos
                    If tipo = "QT" Then
                        QTencontrado = True
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 5)
                                If aux <> "     " Then
                                    If intervalo = 60 Then
                                        totVeh = aux
                                        ini = ini + 6
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![totVeh]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 6
                                        Else
                                            totVeh = acumulado + aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![totVeh]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    
                    'Conteo de vehiculos
                    If tipo = "QP" And QTencontrado = False Then
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 5)
                                If aux <> "     " Then
                                    If intervalo = 60 Then
                                        totVeh = aux
                                        ini = ini + 6
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![totVeh]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 6
                                        Else
                                            totVeh = acumulado + aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![totVeh]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    
                    'Conteo de la velocidad de los vehiculos
                    If tipo = "VT" Then
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 3)
                                If aux <> "   " Then
                                    If intervalo = 60 Then
                                        velMed = aux
                                        ini = ini + 4
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VELMED FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![velMed]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VELMED) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & velMed & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 4
                                        Else
                                            velMed = acumulado + aux
                                            ini = ini + 4
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VELMED FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![velMed]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VELMED) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & velMed & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    'Conteo de vehiculos pesados
                    If tipo = "QL" Then
                        QLencontrado = True
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 5)
                                If aux <> "     " Then
                                    If intervalo = 60 Then
                                        totPes = aux
                                        ini = ini + 6
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![VEHPES]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 6
                                        Else
                                            totVeh = acumulado + aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![VEHPES]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    
                    'Conteo de vehiculos pesados
                    If tipo = "LC" And QLencontrado = False Then
                        limitePeso = Mid(linea, 1, 2)
                        If limitePeso >= 37 Then
                            hora = horaInicio
                            acumulado = -1
                            
                            For i = 1 To numLineasBloque
                                Line Input #2, linea
                                'MsgBox linea
                                longitud = Len(linea)
                                If (longitud <> 78) Then
                                    a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                    progreso.Value = 0
                                    Close #2
                                    Exit Sub
                                End If
                                cont = cont + 1
                                ini = 1
                                For j = 1 To 12
                                    aux = Mid(linea, ini, 5)
                                    If aux <> "     " Then
                                        If intervalo = 60 Then
                                            totPes = aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![VEHPES]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                        End If
                                        If intervalo = 30 Then
                                            If acumulado = -1 Then
                                                acumulado = aux
                                                ini = ini + 6
                                            Else
                                                totVeh = acumulado + aux
                                                ini = ini + 6
                                                '##################
                                                consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Set regAux = .Execute(consulta)
                                                If regAux.EOF = False Then
                                                    If ("" & regAux![VEHPES]) = "" Then
                                                        'MsgBox "hay fila y campo ES NULO"
                                                        'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                        .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                    Else
                                                        'MsgBox "hay fila pero campo NO ES NULO"
                                                        'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                        .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                    End If
                                                Else
                                                    'MsgBox "No hay fila"
                                                    'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                    .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                                End If
                                                '##################
                                                hora = hora + 1
                                                If hora = 24 Then
                                                    hora = 0
                                                    fecha = incrementaFecha(fecha)
                                                End If
                                                acumulado = -1
                                            End If
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    End If
                                        
               
                Else
                    'SE IGNORAN LOS DEMÁS BLOQUES
                
                End If
                

                'cogemos la siguiente linea
                If cont < totalLineas Then
                    Line Input #2, linea
                    longitud = Len(linea)
                    If (longitud <> 78) Then
                        a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                        progreso.Value = 0
                        Close #2
                        Exit Sub
                    End If
                    tipo = Mid(linea, 58, 2)
                End If
                progreso = cont
                cont = cont + 1
                                  
                
            Loop While cont <= totalLineas

            
            .Close
            Close #2
            
        End With
        a = MsgBox("Los datos de '" & sName & "' se importaron correctamente", vbOKOnly, "Importación Correcta")
        
        restab
        
Exit_importarArchivo:
        Exit Sub
Err_importarArchivo:
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Importación")
        progreso.Value = 0
        cnn.Close
        Close #2
        Resume Exit_importarArchivo
End Sub




'######## MÉTODO PARA INCREMENTAR LA FECHA ##########3
Private Function incrementaFecha(ByVal fecha As String) As String
    dia = Mid(fecha, 1, 2)
    mes = Mid(fecha, 4, 2)
    ano = Mid(fecha, 7, 2)

    If (ano Mod 4 = 0) Then
        'bisiesto
        If (mes = 1 Or mes = 3 Or mes = 5 Or mes = 7 Or mes = 8 Or mes = 10 Or mes = 12) Then
            'mes de 31 dias
            If (dia = 31) Then
                dia = 1
                If (mes = 12) Then
                    mes = 1
                    ano = ano + 1
                Else
                    mes = mes + 1
                End If
            Else
                dia = dia + 1
            End If
        Else
            If (mes = 2) Then
                'febrero de 29 dias
                If (dia = 29) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            Else
                'resto meses 30 dias
                If (dia = 30) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            End If
        End If
    Else
        'No bisiesto
        If (mes = 1 Or mes = 3 Or mes = 5 Or mes = 7 Or mes = 8 Or mes = 10 Or mes = 12) Then
            'mes de 31 dias
            If (dia = 31) Then
                dia = 1
                If (mes = 12) Then
                    mes = 1
                    ano = ano + 1
                Else
                    mes = mes + 1
                End If
            Else
                dia = dia + 1
            End If
        Else
            If (mes = 2) Then
                'febrero de 28 dias
                If (dia = 28) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            Else
                'resto meses 30 dias
                If (dia = 30) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            End If
        End If
    End If

    If (Len(dia) = 1) Then
        dia = "0" & dia
    End If
    If (Len(mes) = 1) Then
        mes = "0" & mes
    End If
    If (Len(ano) = 1) Then
        ano = "0" & ano
    End If
    

    incrementaFecha = dia & "/" & mes & "/" & ano
    
    
End Function




'########## MÉTODO DE FILTRADO POR ESTACIÓN ##########3
Private Sub buscar_Click()
    Dim cnn As ADODB.Connection
    Dim SQL As String
    
    progreso.Value = 0
    
    On Error GoTo Err_buscar_Click
                      
        
        ' Establezco la conexión con la base de datos actual
        Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            
            'muestro los datos de la estacion
            consulta = "SELECT IDENTIFICACION FROM EST WHERE IDENTIFICACION='" & estText.Text & "';"
            Set regAux = .Execute(consulta)
            If regAux.EOF <> False Then
                a = MsgBox("La estación a la que hace referencia no está registrada", vbExclamation, "Error de Filtrado")
                restab
                estText.Text = ""
                estText.SetFocus
                Exit Sub
            Else
                consulta = "SELECT IDENTIFICACION, TIPOESTACION, NUMCANALES, CARRETERA, PK, DELEGACION, SITUACION, CANAL1, CANAL2, CANAL3, CANAL4 FROM EST WHERE IDENTIFICACION='" & estText.Text & "';"
                Set regAux = .Execute(consulta)
                estText.Text = "" & regAux![identificacion]
                loc.Visible = True
                locText.Visible = True
                locText.Text = "" & regAux![carretera]
                pk.Visible = True
                pkText.Visible = True
                pkText.Text = "" & regAux![pk]
                tipol.Visible = True
                tipoText.Visible = True
                tipoText.Text = "" & regAux![tipoestacion]
                del.Visible = True
                delText.Visible = True
                delText.Text = "" & regAux![delegacion]
                tra.Visible = True
                traText.Visible = True
                traText.Text = "" & regAux![situacion]
                c1.Visible = True
                c1Text.Visible = True
                c1Text.Text = "" & regAux![canal1]
                c2.Visible = True
                c2Text.Visible = True
                c2Text.Text = "" & regAux![canal2]
                c3.Visible = True
                c3Text.Visible = True
                c3Text.Text = "" & regAux![canal3]
                c4.Visible = True
                c4Text.Visible = True
                c4Text.Text = "" & regAux![canal4]
            End If
            
            

            
            'realizo el filtrado
            .Execute "DELETE * FROM ESTDATAUX"
            If estText.Text = "" Or estText.Text = " " Then
                a = MsgBox("Debe introducir el nombre de una estación", vbExclamation, "Error de Filtrado")
            Else
                .Execute "INSERT INTO ESTDATAUX " & _
                " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
                " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, ((A.VELMED+B.VELMED)/2), A.VEHPES+B.VEHPES " & _
                " FROM ESTDAT AS A, ESTDAT AS B " & _
                " WHERE A.IDEST='" & estText.Text & "' AND A.IDEST= B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
            End If
            AFOROS.RecordSource = "ESTDATAUX"
            a = MsgBox("Los datos se filtrarán", vbOKOnly, "Filtrado")
            'estText.Text = ""
            calendar.Year = Year(Date)
            calendar.Month = Month(Date)
            calendar.Day = Day(Date)
            calendar.Value = fechaAux.Text
            AFOROS.Refresh
            ' Cerramos la conexión
            .Close
            
        End With
        
        'calendar.Value = Today
        
Exit_buscar_Click:
        Exit Sub
Err_buscar_Click:
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Filtrado")
        cnn.Close
        Resume Exit_buscar_Click
End Sub




'####### MÉTODO DE FILTRADO POR FECHA (>= Y NO =) ########
Private Sub calendar_Click()
    Dim cnn As ADODB.Connection
    Dim SQL As String
    Dim fecha As String
    fecha = calendar.Value
    
    dia = Mid(fecha, 1, 2)
    mes = Mid(fecha, 4, 2)
    ano = Mid(fecha, 9, 2)
    
    fecha = dia & "/" & mes & "/" & ano
    'MsgBox fecha
    
    progreso.Value = 0
    
    On Error GoTo Err_calendar_Click
        'construimos la sentencia SQL
        If estText.Text = "" Or estText.Text = " " Then
            SQL = "INSERT INTO ESTDATAUX " & _
                " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
                " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, A.VELMED+B.VELMED, A.VEHPES+B.VEHPES " & _
                " FROM ESTDAT AS A, ESTDAT AS B " & _
                " WHERE A.FECHAMED = '" & fecha & "' AND A.IDEST= B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
        Else
            SQL = "INSERT INTO ESTDATAUX " & _
                " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
                " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, A.VELMED+B.VELMED, A.VEHPES+B.VEHPES " & _
                " FROM ESTDAT AS A, ESTDAT AS B " & _
                " WHERE A.FECHAMED = '" & fecha & "' AND A.IDEST = '" & estText.Text & "' AND A.IDEST = B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
        End If

        ' Establezco la conexión con la base de datos actual
        Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            .Execute "DELETE * FROM ESTDATAUX;"
            .Execute SQL
            AFOROS.RecordSource = "ESTDATAUX"
            a = MsgBox("Los datos se filtrarán", vbOKOnly, "Filtrado")
            AFOROS.Refresh
            ' Cerramos la conexión
            .Close
        End With
        
Exit_calendar_Click:
        Exit Sub
Err_calendar_Click:
        cnn.Close
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Filtrado")
        Resume Exit_calendar_Click
End Sub




'########## CON ESTOS MÉTODOS MOSTRAMOS LA INFORMACIÓN DE LA BASE DE DATOS
'DE NUESTRA APLICACION ###################################################
Private Sub info_Click()
    progreso.Value = 0
    examinarArchivo (App.Path & "/BaseDatos.mdb")
End Sub


Sub examinarArchivo(nomArchivo As String)
    On Error GoTo ErrorProducido
    Dim fso As Object 'New FileSystemObject
    Dim archivo As Object 'File
    Dim mensaje As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.GetFile(nomArchivo)
    
    'mensaje = vbCrLf & "Fecha creación: " _
        & archivo.DateCreated _
        & vbCrLf & "Fecha último acceso: " _
        & archivo.DateLastAccessed _
        & vbCrLf & "Fecha última modificación: " _
        & archivo.DateLastModified _
        & vbCrLf & "Letra unidad: " _
        & archivo.Drive _
        & vbCrLf & "Nombre archivo: " _
        & archivo.Name _
        & vbCrLf & "Carpeta: " _
        & archivo.ParentFolder _
        & vbCrLf & "Ruta completa: " _
        & archivo.Path _
        & vbCrLf & "Nombre archivo DOS: " _
        & archivo.ShortName _
        & vbCrLf & "Nombre ruta DOS: " _
        & archivo.ShortPath _
        & vbCrLf & "Tamaño: " _
        & (Format(archivo.Size, "##,###")) / 1024 & " KB" _
        & vbCrLf & "Tipo de archivo: " _
        & archivo.Type
        
    Dim tam As Double
    tam = (Format(archivo.Size, "##,###")) / 1024
    tam = tam / 1024
    
    
    mensaje = vbCrLf & "Fecha creación:" & vbTab & vbTab & "" _
        & archivo.DateCreated _
        & vbCrLf & "Fecha último acceso:" & vbTab & "" _
        & archivo.DateLastAccessed _
        & vbCrLf & "Fecha última modificación:" & vbTab & "" _
        & archivo.DateLastModified _
        & vbCrLf & vbCrLf & "Ruta completa:  " _
        & archivo.Path _
        & vbCrLf & vbCrLf & "Tamaño:  " _
        & tam & "  MB" _
        & vbCrLf & vbCrLf & "Tipo de archivo:  " _
        & archivo.Type
        
    Set archivo = Nothing
    Set fso = Nothing
    
    a = MsgBox(vbCrLf & "INFORMACIÓN:" & vbCrLf & mensaje, vbInformation, "INFORMACIÓN DE LA BASE DE DATOS")
Salida:
    Exit Sub
ErrorProducido:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al recopilar Información")
    Resume Salida
End Sub





'##### CON ESTOS DOS MÉTODOS SE RESTABLECEN LOS DATOS DE LA VISTA DE LA APLICACION
'QUE MUESTRA LOS DATOS DE LA BASE DE DATOS. SE USA DESPUÉS DE UN FILTRADO (MANUALMENTE),
'O AUTOMATICAMENTE AL INICIAR LA APLICACION ####################################

Private Sub restablecer_Click()
    a = MsgBox("Se restablecerán los datos", vbOKOnly, "Restablecer Datos")
    restab
End Sub



Private Sub restab()
    calendar.Year = Year(Date)
    calendar.Month = Month(Date)
    calendar.Day = Day(Date)
    calendar.Value = fechaAux.Text
    
    AFOROS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseDatos.mdb;Persist Security Info=False"
    AFOROS.RecordSource = "ESTDATAUX"
    Dim cnn As ADODB.Connection
    On Error GoTo err_restablecer
    Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            .Execute "DELETE * FROM ESTDATAUX;"
            SQL = "INSERT INTO ESTDATAUX " & _
            " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
            " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, ((A.VELMED+B.VELMED)/2), A.VEHPES+B.VEHPES " & _
            " FROM ESTDAT AS A, ESTDAT AS B " & _
            " WHERE A.IDEST= B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
            .Execute SQL
            .Close
        End With
    AFOROS.Refresh
    
    progreso.Value = 0
    
    estText.Text = ""
    loc.Visible = False
    locText.Visible = False
    pk.Visible = False
    pkText.Visible = False
    tipol.Visible = False
    tipoText.Visible = False
    del.Visible = False
    delText.Visible = False
    tra.Visible = False
    traText.Visible = False
    c1.Visible = False
    c1Text.Visible = False
    c2.Visible = False
    c2Text.Visible = False
    c3.Visible = False
    c3Text.Visible = False
    c4.Visible = False
    c4Text.Visible = False
    
salir:
    Exit Sub
err_restablecer:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al Restablecer")
    cnn.Close
    Resume salir
End Sub




'######### MÉTODO PARA SALIR DE LA APLICACION #####################
Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub
