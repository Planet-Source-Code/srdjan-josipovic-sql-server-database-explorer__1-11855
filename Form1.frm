VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "SQL Database Explorer"
   ClientHeight    =   10920
   ClientLeft      =   -180
   ClientTop       =   495
   ClientWidth     =   15420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   15420
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRtbox 
      Caption         =   "Close"
      Height          =   375
      Left            =   13200
      TabIndex        =   26
      Top             =   10200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEmng 
      Caption         =   "Enterprise Manager"
      Height          =   975
      Left            =   2760
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":1594
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   975
      Left            =   1440
      Picture         =   "Form1.frx":19D6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame fraExplorer 
      Caption         =   "Explorer"
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2640
         Top             =   2640
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2818
               Key             =   "imgDatabase"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2A64
               Key             =   "imgField"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2B76
               Key             =   "imgProp"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   7695
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   13573
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   45
      End
   End
   Begin RichTextLib.RichTextBox rtBox 
      Height          =   8700
      Left            =   8640
      TabIndex        =   25
      Top             =   1515
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   15346
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":2C8E
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Query "
      Height          =   8775
      Left            =   4800
      TabIndex        =   66
      Top             =   1440
      Visible         =   0   'False
      Width           =   9495
      Begin VB.OptionButton optQuery 
         Caption         =   "Select Query"
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDML 
         Caption         =   "Execute DML"
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   360
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   5535
         Left            =   0
         TabIndex        =   72
         Top             =   3240
         Visible         =   0   'False
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9763
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
         Height          =   735
         Left            =   7440
         Picture         =   "Form1.frx":2D63
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   735
         Left            =   8400
         TabIndex        =   70
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute (DML)"
         Enabled         =   0   'False
         Height          =   735
         Left            =   6480
         Picture         =   "Form1.frx":2EAD
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   360
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5535
         Left            =   0
         TabIndex        =   68
         Top             =   3240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9763
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtQuery 
         Height          =   1695
         Left            =   120
         TabIndex        =   69
         Top             =   1320
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2990
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   1
         TextRTF         =   $"Form1.frx":2FF7
      End
   End
   Begin VB.Frame fraEmng 
      Caption         =   "Enterprise Manager"
      Height          =   5295
      Left            =   4800
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Frame fraWizards 
         Caption         =   "EM Wizards :"
         Height          =   975
         Left            =   3360
         TabIndex        =   62
         Top             =   240
         Width           =   4815
         Begin VB.CommandButton cmdDTSEx 
            Caption         =   "DTS Export"
            Height          =   375
            Left            =   3240
            TabIndex        =   65
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdDTSIm 
            Caption         =   "DTS Import"
            Height          =   375
            Left            =   1680
            TabIndex        =   64
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdWizard 
            Caption         =   "DB Maintenance "
            Height          =   375
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraSub2 
         Caption         =   "Administration :"
         Height          =   3375
         Left            =   3360
         TabIndex        =   39
         Top             =   1680
         Width           =   2295
         Begin VB.CommandButton cmdProperties 
            Caption         =   "Table Properties"
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdPermission 
            Caption         =   "Permissions"
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdDependencies 
            Caption         =   "Dependencies"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CommandButton cmdIndexes 
            Caption         =   "Manage Indexes"
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdScripts 
            Caption         =   "Generate Scripts"
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   2760
            Width           =   1815
         End
      End
      Begin VB.Frame fraSub1 
         Caption         =   "Implementation :"
         Height          =   3375
         Left            =   720
         TabIndex        =   40
         Top             =   1680
         Width           =   2295
         Begin VB.CommandButton cmdNewRule 
            Caption         =   "New Rule"
            Height          =   375
            Left            =   240
            TabIndex        =   60
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CommandButton cmdNewUdt 
            Caption         =   "New UDT"
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CommandButton cmdNewDef 
            Caption         =   "New Default"
            Height          =   375
            Left            =   240
            TabIndex        =   58
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton cmdNewSp 
            Caption         =   "New Stored Procedure"
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdNewTrigg 
            Caption         =   "New Trigger"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton CmdNewRole 
            Caption         =   "New Role"
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame fraSub3 
         Caption         =   "DB Administration :"
         Height          =   3375
         Left            =   6000
         TabIndex        =   33
         Top             =   1680
         Width           =   2175
         Begin VB.CommandButton cmdNewdbUser 
            Caption         =   "New DB User"
            Height          =   375
            Left            =   240
            TabIndex        =   61
            Top             =   2760
            Width           =   1695
         End
         Begin VB.CommandButton cmddbProperties 
            Caption         =   "Database Properties"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdSecurity 
            Caption         =   "SQL Security"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CommandButton cmdBackup 
            Caption         =   "Backup Database"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton cmdRestore 
            Caption         =   "Restore Database"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton cmdShrink 
            Caption         =   "Shrink Database"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdHideNS 
         Caption         =   "Hide"
         Height          =   375
         Left            =   8640
         TabIndex        =   30
         Top             =   4680
         Width           =   495
      End
      Begin VB.ComboBox cmbTablesNS 
         Height          =   315
         Left            =   720
         TabIndex        =   29
         Text            =   "Tables"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image imgEmng 
         Height          =   495
         Left            =   8640
         Picture         =   "Form1.frx":30CC
         Stretch         =   -1  'True
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search"
      Height          =   5415
      Left            =   4800
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.ListBox lstResult 
         Height          =   2010
         Left            =   480
         TabIndex        =   7
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Top             =   1920
         Width           =   3375
      End
      Begin VB.ComboBox cmbColumns 
         Height          =   315
         Left            =   2400
         TabIndex        =   5
         Text            =   "Columns"
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox cmbTables 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "Tables"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Image imgSearch 
         Height          =   375
         Left            =   2160
         Picture         =   "Form1.frx":3D96
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblRecords 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   4920
         Width           =   45
      End
      Begin VB.Label lblResult 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label lblCriteria 
         Caption         =   "Search Text:"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Table Name and Column Name :"
         Height          =   195
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   2805
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   8775
      Left            =   4800
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdConstrScroll 
         Caption         =   "Constraints >>"
         Height          =   375
         Left            =   2400
         TabIndex        =   55
         Top             =   6240
         Width           =   1215
      End
      Begin VB.ListBox lstIdConstraint 
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   6240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox lstConstraint 
         Height          =   645
         Left            =   240
         TabIndex        =   53
         Top             =   5520
         Width           =   3375
      End
      Begin VB.CommandButton cmdDefScroll 
         Caption         =   "Defaults >>"
         Height          =   375
         Left            =   2400
         TabIndex        =   52
         Top             =   4800
         Width           =   1215
      End
      Begin VB.ListBox lstIdDef 
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   4800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox lstDefault 
         Height          =   645
         Left            =   240
         TabIndex        =   50
         Top             =   4080
         Width           =   3375
      End
      Begin VB.ComboBox cmbFK 
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Text            =   "Foreign Key"
         Top             =   7440
         Width           =   3375
      End
      Begin VB.ComboBox cmbPK 
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Text            =   "Primary Key"
         Top             =   6840
         Width           =   3375
      End
      Begin VB.CommandButton cmdHideinfo 
         Caption         =   "Hide"
         Height          =   375
         Left            =   3120
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdScrollTrig 
         Caption         =   "Triggers >>"
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ListBox lstTrigg 
         Height          =   645
         Left            =   240
         TabIndex        =   18
         Top             =   2640
         Width           =   3375
      End
      Begin VB.ListBox lstId 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbView 
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Text            =   "Views"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.ComboBox cmbStoredproc 
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Text            =   "Stored Procedures"
         Top             =   840
         Width           =   3375
      End
      Begin VB.Image imgInfo 
         Height          =   255
         Left            =   1800
         Picture         =   "Form1.frx":4BD8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblConstraint 
         Caption         =   "Constraints :"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label lblDefaults 
         Caption         =   "Defaults :"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblTrigger 
         Caption         =   "Triggers :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblView 
         Caption         =   "Views :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblStProc 
         Caption         =   "Stored Procedures :"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *****************************************************
' * Application : # SQL Server Database Explorer #   **
' *           by Srdjan Josipovic                    **
' *           Belgrade/Yugoslavia                    **                      
' *           srdjan.j@sezampro.yu                   **
' *     http://vbdeveloper.webjump.com               **                      
' *   Copyright © 2000 ; Srdjan Josipovic            **     
' *          All Rights Reserved                     **                       
' *****************************************************


' NOTE , this is just an extract of whole application , but I hope that You can find it useful
' for your own projects.




Option Explicit

Dim user As String                       ' User Name
Dim Pass As String                       ' Password
Dim srvName As String                    ' Server Name
Dim dbName As String                     ' Database Name
Dim rsCol As ADODB.Recordset             ' Recordset - Execute searching Criteria and return data
Dim rs As ADODB.Recordset                ' Recordset - basic recset filling up the objects in Form_load event
Dim rsTree As ADODB.Recordset            ' Recordset - Complex Inner join and similar SQL queries
Dim rsTView As ADODB.Recordset           ' Recordset - In Click event for Treeview , exploring db (db>tables>columns)
Dim rsQuery As ADODB.Recordset           ' Recordset - Execution of Select Statement
Dim WithEvents cn As ADODB.Connection    ' Connection - ADO connection object , For Insert , Update, Delete ( DML Commands )
Attribute cn.VB_VarHelpID = -1
Dim cmd As Command                       ' Command - execution within Connection
Dim Column As Field                      ' Table Columns
Dim strCol As String                     ' Retreive Column names from table and fill the combo with them
Dim Query As String                      ' DataSource for opening recordsets
Dim strFind As String                    ' Complex SQL Query - in Searching
Dim Key As String                        ' Searching Criteria
Dim queryTree As String                  ' SQL Query - variable
Dim objSQLNS As SQLNamespace             ' SQL Namespace
Dim objSQLNSObj As SQLNamespaceObject    ' SQL Namespace Object
Dim srvSQL As SQLServer                  ' SQLServer , part of SQL DMO Object
Dim hArray(10) As Long                   ' Initializing SQL Namespace Objects
Dim tNode As Node                        ' Treeview node
Dim i As Integer
Dim nodetext As String                   ' Variable for filling up Treeview , with db records
Dim tvNode As Node                       ' Treeview node
Dim tvTag                                ' Tag , for exploring Treeview
Dim tvText As String





Private Sub cmbStoredproc_Click()

On Error GoTo eh_sqldmo
    'Initialize SQL DMo Object
    Set srvSQL = New SQLServer
    'connect to server db with SQL DMO Object
    srvSQL.Connect srvName, user, Pass
    rtBox.Visible = True
    cmdRtbox.Visible = True
    'set text property of rtbox to Stored Procedure text
    rtBox.Text = srvSQL.Databases(dbName).StoredProcedures(cmbStoredproc.Text).Text
    srvSQL.Disconnect
    Set srvSQL = Nothing
    fraSearch.Visible = False
    
    Exit Sub
eh_sqldmo:
  MsgBox Err.Description, vbCritical, "Error Initializing SQL DMO Object"
End Sub

Private Sub cmbTables_Click()
On Error GoTo eh_col

 
   Set rsCol = New ADODB.Recordset
   ' fill the second combo with column names from table
   strCol = "Select * from " & "[" & cmbTables.Text & "]"
   rsCol.Open strCol, "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & "", adOpenStatic, adLockReadOnly
       cmbColumns.Clear
       
     For Each Column In rsCol.Fields
       cmbColumns.AddItem Column.Name
     Next
       cmbColumns.Text = "Columns"
   rsCol.Close
Exit Sub
eh_col:

MsgBox Err.Description, vbCritical, "Error"

End Sub



Private Sub cmbView_Click()
On Error GoTo eh_sqldmo


   
    'Initializing SQL DMo Object
    Set srvSQL = New SQLServer
    srvSQL.Connect srvName, user, Pass
    rtBox.Visible = True
    cmdRtbox.Visible = True
    'Set property of rtbox to text of selected View
    rtBox.Text = srvSQL.Databases(dbName).Views(cmbView.Text).Text
    srvSQL.Disconnect
    Set srvSQL = Nothing
    fraSearch.Visible = False
     
     Exit Sub
eh_sqldmo:
     MsgBox Err.Description, vbCritical, "Error Initializing SQL DMO Object"
End Sub

Private Sub cmdBackup_Click()
On Error GoTo ErrHandler
    
    ' Get first level server->databases.
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    ' Get second level server->databases->database
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    'Initialize SQL Namespace Object
    ' Get a SQLNamespaceObject object to execute commands against at the desired level.
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    'Execute command for M$ Console Management
    objSQLNSObj.Commands("Backup Database").Execute

Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdClear_Click()
    txtQuery.Text = ""
End Sub

Private Sub cmdConstrScroll_Click()
On Error Resume Next
  'Check for existing Constraints
  If lstIdConstraint.ListCount = 0 Then
     MsgBox "No Constraint available !", vbInformation, "Constraint"
     Exit Sub
  Else
  End If
    'Syntax for scrolling Constraints and automatically see the value
    'in the rtbox
    If lstIdConstraint.ListIndex = lstIdConstraint.ListCount - 1 Then
        lstIdConstraint.ListIndex = 0
    Else
        lstIdConstraint.ListIndex = (lstIdConstraint.ListIndex + 1)
    End If
    'lstID visible property is set to false , cause it is populate with
    'ID number , and there is no need for showing it to the user
    lstConstraint.ListIndex = lstIdConstraint.ListIndex
    rtBox.Visible = True
    cmdRtbox.Visible = True
    'recordset which return this complex query , and text property of existing Constraints
    rsTree.Open "SELECT o.id, c.text FROM sysobjects o INNER JOIN syscomments c ON o.id = c.id WHERE o.type = 'C'"
    rsTree.MoveFirst
    'Fundamental syntax for scrolling
    rsTree.Find "id =" & lstIdConstraint.List(lstIdConstraint.ListIndex)
    rtBox.Text = rsTree.Fields("text")
    rsTree.Close
 
End Sub

Private Sub cmddbProperties_Click()
  
  On Error GoTo ErrHandler
    'Same syntax as above for level hArray(2)
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    objSQLNSObj.Commands("Properties").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdDefScroll_Click()
On Error Resume Next
   If lstIdDef.ListCount = 0 Then
      MsgBox "No Defaults available !", vbInformation, "Defaults"
   Exit Sub
   Else
   End If
   
   If lstIdDef.ListIndex = lstIdDef.ListCount - 1 Then
        lstIdDef.ListIndex = 0
    Else
        lstIdDef.ListIndex = (lstIdDef.ListIndex + 1)
   End If
    
    lstDefault.ListIndex = lstIdDef.ListIndex
    rtBox.Visible = True
    cmdRtbox.Visible = True
    'same as above for Constraints
    rsTree.Open "SELECT o.id, c.text FROM sysobjects o INNER JOIN syscomments c ON o.id = c.id WHERE o.type = 'D'"
    rsTree.MoveFirst
    rsTree.Find "id =" & lstIdDef.List(lstIdDef.ListIndex)
    rtBox.Text = rsTree.Fields("text")
    rsTree.Close
 
End Sub

Private Sub cmdDelete_Click()
   
  If cmbTablesNS.Text = "Tables" Then
    MsgBox "Table Name is missing !", vbInformation, "Delete"
    Exit Sub
  ElseIf cmbTables.Text = "" Then
    MsgBox "Table Name is missing !", vbInformation, "Delete"
  End If
  
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    'get thirrd level server->databases->database->tables
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 24) ' Please , for those constants see Word documents attached
    'Get 4th level server->databases->database->tables->table                constant 24-SQLNSOBJECTTYPE_TABLES
    hArray(4) = objSQLNS.GetFirstChildItem(hArray(3), 25, cmbTablesNS.Text) 'constant 25-SQLNSOBJECTTYPE_TABLE and Table name from Combo box
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(4))
    objSQLNSObj.Commands("Delete").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdDependencies_Click()
   
  If cmbTablesNS.Text = "Tables" Then
    MsgBox "Table Name is missing !", vbInformation, "Dependencies"
    Exit Sub
  ElseIf cmbTables.Text = "" Then
     MsgBox "Table Name is missing !", vbInformation, "Dependencies"
    End If
  
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 24)
    hArray(4) = objSQLNS.GetFirstChildItem(hArray(3), 25, cmbTablesNS.Text)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(4))
    objSQLNSObj.Commands("Object Dependencies").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdDTSEx_Click()
On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    objSQLNSObj.Commands("Data Export").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub


Private Sub cmdDTSIm_Click()
On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    objSQLNSObj.Commands("Data Import").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub


Private Sub cmdEmng_Click()
  fraEmng.Visible = True
  fraSearch.Visible = False
  fraInfo.Visible = False
  rtBox.Visible = False
  cmdRtbox.Visible = False
  fraGrid.Visible = False
End Sub

Private Sub cmdExecute_Click()
   
   
   'Execute standard and complex SQL Select Query
   ' You can also execute queries like:
   ' SELECT City , Count (CompanyName) FROM Customers GROUP BY City
   ' or
   'SELECT ORDERID,SUM(UnitPrice)AS Total FROM [Order Details]GROUP BY ORDERID
   ' or something like this
   'SELECT CompanyName,City FROM Customers WHERE City LIKE 'L%' AND  City LIKE '%ND%'
   ' or even this
   'SUM (UnitPrice*Quantity) AS TOTAL
   'From [Order Details]
   'WHERE OrderID BETWEEN 10500 AND 10600
   'GROUP BY OrderID
   'Having Sum(UnitPrice * Quantity) > 3000

   ' Note : Use Northwind for those Queries , otherwise they would not work !
   
   ' It will even create whole table ,and You can use pure T-SQL : see this example
   '    CREATE TABLE Buyers
   '    (
   '    IDBuyer decimal (18,0) IDENTITY (100,2) Not Null,
   '    Name nvarchar(40) Not Null,
   '    City nvarchar(30) Null DEFAULT 'Beograd',
   '    Date datetime Not Null DEFAULT GetDate(),
   '    Category  VarChar(6) Not Null
   '    CONSTRAINT catChk CHECK
   '    (Category LIKE '[A-C][0-9]-[0-9][0-9][0-9]'),
   '    Description ntext Null,
   '    CONSTRAINT BuyersPK PRIMARY KEY (IDBuyer,Name),
   '    CONSTRAINT DateChk CHECK (Date  >= GetDate())
   '    )

   
   ' However ,it supports DROP , ALTER , and I think all other T-SQL Commands
   '
   
   'You can also perform something like this :
   
   '   DECLARE @SC money
   '    Select @SC=AVG(UnitPrice) From Products
   '    Select @SC AS [Srednja cena]
   '    Update Products
   '    Set UnitPrice = UnitPrice*
   '     Case
   '       When UnitPrice > @SC Then 0.95
   '       When UnitPrice < @SC Then 1.05
   '    Else 1.0
   ' End

   ' You can create Stored Procedure : simple one
   
   '  CREATE PROC Sum
   '  @a tinyint ,@b tinyint ,@c tinyint OUTPUT
   '  AS
   '  SET @c=@a+@b


   ' Same thing is with Triggers and Views
   
   ' It is not like Query Analyzer , but I find this useful when I am not able to work with
   ' Enterprise Manager or Query Analyzer
   
   
   
   
   
   
   ' ADO connection for Update , Insert , Delete commands
   If Trim(txtQuery) = "" Then Exit Sub

On Error GoTo e_TrapCon

   Set cn = New Connection
   cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & ""
   cn.Open
   'Check if cn is open
   Debug.Assert (cn.State = adStateOpen)
   Query = txtQuery.Text
   
   Set cmd = New Command
   Set cmd.ActiveConnection = cn
   cmd.CommandText = Query
   cmd.CommandTimeout = 15
   cmd.Execute
   Call TreeView1_Click
   DataGrid1.Visible = True
   DataGrid2.Visible = False

   MsgBox "Execution Complete !", vbInformation, "Execute"
   cn.Close
   Set cn = Nothing
Exit Sub
e_TrapCon:
    MsgBox "Error: " & Err.Description, vbCritical, "Execution Failed !"
End Sub

Private Sub cmdHide_Click()
  fraSearch.Visible = False
End Sub

Private Sub cmdHideinfo_Click()
  fraInfo.Visible = False
End Sub

Private Sub cmdHideNS_Click()
  fraEmng.Visible = False
End Sub

Private Sub cmdIndexes_Click()
  
  If cmbTablesNS.Text = "Tables" Then
    MsgBox "Table Name is missing !", vbInformation, "Manage Indexes"
    Exit Sub
  ElseIf cmbTables.Text = "" Then
     MsgBox "Table Name is missing !", vbInformation, "Manage Indexes"
    End If
  
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 24)
    hArray(4) = objSQLNS.GetFirstChildItem(hArray(3), 25, cmbTablesNS.Text)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(4))
    objSQLNSObj.Commands("Manage Indexes").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdInfo_Click()
   fraInfo.Visible = True
   fraSearch.Visible = False
   fraEmng.Visible = False
   fraGrid.Visible = False
End Sub

Private Sub cmdNewdbUser_Click()
On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 18) ' constant 18 is SQLNSOBJECTTYPE_DATABASE_USERS
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(3))
    objSQLNSObj.Commands("New Database user").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub


Private Sub cmdNewDef_Click()
On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 34) 'constant 34 is SQLNSOBJECTTYPE_DATABASE_DEFAULTS
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(3))
    objSQLNSObj.Commands("New Default").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub


Private Sub cmdNewRule_Click()
On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 32) 'constant 32 is SQLNSOBJECTTYPE_DATABASE_RULES
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(3))
    objSQLNSObj.Commands("New Rule").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub


Private Sub cmdNewSp_Click()
  
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 28) 'constant 28 is SQLNSOBJECTTYPE_DATABASE_SPS
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(3))
    objSQLNSObj.Commands("New Stored Procedure").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdNewTrigg_Click()
   
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 24)
    hArray(4) = objSQLNS.GetFirstChildItem(hArray(3), 25)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(4))
    objSQLNSObj.Commands("Manage Triggers").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub CmdNewRole_Click()
On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 20) 'constan 20 is - SQLNSOBJECTTYPE_DATABASE_ROLES
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(3))
    objSQLNSObj.Commands("New Database Role").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdNewUdt_Click()
On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 36) 'constant 36 is - SQLNSOBJECTTYPE_DATABASE_UDDTS
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(3))
    objSQLNSObj.Commands("New User Defined Data Type").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub


Private Sub cmdPermission_Click()
   
  If cmbTablesNS.Text = "Tables" Then
    MsgBox "Table Name is missing !", vbInformation, "Object Permissions"
    Exit Sub
  ElseIf cmbTables.Text = "" Then
     MsgBox "Table Name is missing !", vbInformation, "Object Permissions"
    End If
  
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 24)
    hArray(4) = objSQLNS.GetFirstChildItem(hArray(3), 25, cmbTablesNS.Text)
 
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(4))
    objSQLNSObj.Commands("Object Permissions").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdQuery_Click()
  If Trim(txtQuery) = "" Then Exit Sub
  On Error GoTo e_Trap
        
       
    Query = txtQuery.Text
    Set rsQuery = New ADODB.Recordset
       rsQuery.Open Query, "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & "", adOpenStatic, adLockReadOnly
    Set DataGrid2.DataSource = rsQuery
       DataGrid1.Visible = False
       DataGrid2.Visible = True
       
Exit Sub
e_Trap:
    MsgBox "Error: " & Err.Description, vbCritical, "Execution Failed !"
End Sub

Private Sub cmdRestore_Click()
On Error GoTo ErrHandler

    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    objSQLNSObj.Commands("Restore Database").Execute

Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdRtbox_Click()
  rtBox.Visible = False
  cmdRtbox.Visible = False
End Sub

Private Sub cmdScripts_Click()
     
  If cmbTablesNS.Text = "Tables" Then
     MsgBox "Table Name is missing !", vbInformation, "Generate Scripts"
     Exit Sub
  ElseIf cmbTables.Text = "" Then
     MsgBox "Table Name is missing !", vbInformation, "Generate Scripts"
  End If
  
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 24)
    hArray(4) = objSQLNS.GetFirstChildItem(hArray(3), 25, cmbTablesNS.Text)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(4))
    objSQLNSObj.Commands("Generate Scripts").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdScrollTrig_Click()
On Error Resume Next
 If lstId.ListCount = 0 Then
    MsgBox "No Triggers available !", vbInformation, "Triggers"
    Exit Sub
  Else
  End If

  If lstId.ListIndex = lstId.ListCount - 1 Then
      lstId.ListIndex = 0
  Else
      lstId.ListIndex = (lstId.ListIndex + 1)
  End If
 
      lstTrigg.ListIndex = lstId.ListIndex
      rtBox.Visible = True
      cmdRtbox.Visible = True
      'SQL Inner Join query to get all triggers and their text property
      rsTree.Open "SELECT o.id, c.text FROM sysobjects o INNER JOIN syscomments c ON o.id = c.id WHERE o.type = 'tr'"
        rsTree.MoveFirst
        rsTree.Find "id =" & lstId.List(lstId.ListIndex)
        rtBox.Text = rsTree.Fields("text")
      rsTree.Close
 
End Sub

Private Sub cmdSearch_Click()
   fraSearch.Visible = True
   fraInfo.Visible = False
   fraEmng.Visible = False
   rtBox.Visible = False
   cmdRtbox.Visible = False
   fraGrid.Visible = False
End Sub

Private Sub cmdProperties_Click()
   
  If cmbTablesNS.Text = "Tables" Then
    MsgBox "Table Name is missing !", vbInformation, "Table Properties"
    Exit Sub
  ElseIf cmbTables.Text = "" Then
     MsgBox "Table Name is missing !", vbInformation, "Table Properties"
  End If
  
  On Error GoTo ErrHandler
  
    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    hArray(3) = objSQLNS.GetFirstChildItem(hArray(2), 24)
    hArray(4) = objSQLNS.GetFirstChildItem(hArray(3), 25, cmbTablesNS.Text)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(4))
    objSQLNSObj.Commands("Properties").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdSecurity_Click()
On Error GoTo ErrHandler

    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    objSQLNSObj.Commands("Manage SQL Server Security").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdShrink_Click()
On Error GoTo ErrHandler

    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    objSQLNSObj.Commands("Shrink Database").Execute

Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub

Private Sub cmdWizard_Click()
On Error GoTo ErrHandler

    hArray(1) = objSQLNS.GetFirstChildItem(hArray(0), SQLNSOBJECTTYPE_DATABASES)
    hArray(2) = objSQLNS.GetFirstChildItem(hArray(1), SQLNSOBJECTTYPE_DATABASE, dbName)
    Set objSQLNSObj = objSQLNS.GetSQLNamespaceObject(hArray(2))
    objSQLNSObj.Commands("Database Maintenance Plan").Execute
Cleanup:
    Set objSQLNSObj = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description & " " & Err.Number, vbOKOnly, "Error"
    GoTo Cleanup
End Sub


Private Sub Form_Load()
On Error GoTo EH

  user = frmLogin!txtUser
  Pass = frmLogin!txtPass
  srvName = frmLogin!txtServer
  dbName = frmLogin!txtDatabase

  Set rsTree = New ADODB.Recordset
  Set rs = New ADODB.Recordset
  Set tNode = TreeView1.Nodes.Add(, , , "Server : " & srvName)
  Set tNode = TreeView1.Nodes.Add(, , "M", dbName, "imgDatabase")
    ' Select all User Tables from selected database
    Query = " Select * from sysobjects where xtype='U'"
    rs.Open Query, "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & "", adOpenStatic, adLockReadOnly

  ' Populate TreeView with table names ,
  Do While Not rs.EOF
     nodetext = rs.Fields("name")
     Set tNode = TreeView1.Nodes.Add("M", tvwChild, nodetext, nodetext, "imgField")
       tNode.Tag = "Tables"
       cmbTables.AddItem rs.Fields("name")
       cmbTablesNS.AddItem rs.Fields("name")
       tNode.EnsureVisible
       queryTree = "Select * From " & "[" & nodetext & "]"
          rsTree.Open queryTree, "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & "", adOpenStatic, adLockReadOnly
            For Each Column In rsTree.Fields
              Set tNode = TreeView1.Nodes.Add(nodetext, tvwChild, , Column.Name, "imgProp")
              tNode.Tag = "Fields"
              tNode.Expanded = False
            Next
          rsTree.Close
          rs.MoveNext
  Loop

       lblInfo.Caption = "Exploring Database :" & dbName & "/" & " on Server :" & srvName
       rs.Close
'######### Stored Procedures
  rs.Open "Select * From sysobjects where xtype = 'p'"
  Do While Not rs.EOF
     cmbStoredproc.AddItem rs.Fields("name")
     rs.MoveNext
  Loop
  rs.Close
'########## Views
  rs.Open "Select * From sysobjects where xtype='v'"
  Do While Not rs.EOF
    cmbView.AddItem rs.Fields("name")
    rs.MoveNext
  Loop
  rs.Close
'####### Triggers
' Here you can see that I have been working with 2 rsets as one will return id and text ,and secondone - name
' of the trigger , I will use SQL DMO to populate rtbox with Trigger's text
  rs.Open "Select * from sysobjects where xtype='TR'"
  rsTree.Open "SELECT o.id, c.text FROM sysobjects o INNER JOIN syscomments c ON o.id = c.id WHERE o.type = 'tr'"

  Do While Not rs.EOF
    lstTrigg.AddItem rs.Fields("name")
    rs.MoveNext
  Loop
  rs.Close

  Do While Not rsTree.EOF
    lstId.AddItem rsTree.Fields("id")
    rsTree.MoveNext
  Loop
  rsTree.Close

'#####Constraint

  rs.Open "Select * from sysobjects where xtype='C'"
  rsTree.Open "SELECT o.id, c.text FROM sysobjects o INNER JOIN syscomments c ON o.id = c.id WHERE o.type = 'C'"
  Do While Not rs.EOF
     lstConstraint.AddItem rs.Fields("name")
     rs.MoveNext
  Loop
  rs.Close

  Do While Not rsTree.EOF
    lstIdConstraint.AddItem rsTree.Fields("id")
    rsTree.MoveNext
  Loop
  rsTree.Close

'#########Default
  rs.Open "Select * From sysobjects where xtype='D'"
  rsTree.Open "SELECT o.id, c.text FROM sysobjects o INNER JOIN syscomments c ON o.id = c.id WHERE o.type = 'D'"

  Do While Not rs.EOF
    lstDefault.AddItem rs.Fields("name")
     rs.MoveNext
  Loop
  rs.Close

  Do While Not rsTree.EOF
    lstIdDef.AddItem rsTree.Fields("id")
    rsTree.MoveNext
  Loop
  rsTree.Close


'##########Primary Key

' This will populate Combo box with names of Primary keys :
' Note , PK - Primary Key
'        UPK - Unique Primary Key
'        UPKCL - Unique Primary Key Clustered
  rs.Open "Select * From sysobjects Where xtype='PK'"
  Do While Not rs.EOF
    cmbPK.AddItem rs.Fields("name")
    rs.MoveNext
  Loop
  rs.Close
'############ Foreign Key

  rs.Open "select * from sysobjects where xtype='F'"
  Do While Not rs.EOF
    cmbFK.AddItem rs.Fields("name")
    rs.MoveNext
  Loop
  rs.Close

' Initialize SQL NameSpace Object
  Set objSQLNS = New SQLNamespace

  objSQLNS.Initialize "SQL Server Database Explorer", SQLNSRootType_Server, "Server=" & srvName & ";UID=" & user & ";pwd=" & Pass & ";", hWnd
' Get a root object
  hArray(0) = objSQLNS.GetRootItem

Exit Sub
EH:
  Unload Me
  frmLogin.Show

  MsgBox "Please , Check Your User ID,Password,Server Name and Database Name !", vbCritical, "Login Failed"
Exit Sub
End Sub

Private Sub Form_Terminate()
   Set objSQLNS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Unload everything from memory ,that you will be able to connect to other db within same connection
   ' in the other case you will get error that SQL Namespacwe object is already initialized.
   Set frmMain = Nothing
   frmLogin.Show
End Sub

Private Sub optDML_Click()
   cmdQuery.Enabled = False
   cmdExecute.Enabled = True
   txtQuery = ""
End Sub

Private Sub optQuery_Click()
   cmdExecute.Enabled = False
   cmdQuery.Enabled = True
   txtQuery = ""
End Sub

Private Sub TreeView1_Click()
   On Error GoTo eh_tvw
   
   Set tvNode = TreeView1.SelectedItem
   tvText = tvNode.Text
   tvTag = tvNode.Tag
  'Exploring all tables and columns in treeview
  'and also show the records in the data grid of selected node
   Set rsTView = New ADODB.Recordset
     If tvTag = "Tables" Then
  
        rsTView.Open "Select * from " & "[" & tvText & "]", "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & "", adOpenStatic, adLockReadOnly
        Set DataGrid1.DataSource = rsTView
        DataGrid2.Visible = False
        DataGrid1.Visible = True
        fraGrid.Visible = True
        fraEmng.Visible = False
        fraSearch.Visible = False
        fraInfo.Visible = False
        rtBox.Visible = False
        cmdRtbox.Visible = False
     Else
End If
    
    If tvTag = "Fields" Then
      rsTView.Open "Select " & "[" & tvText & "]" & " From " & "[" & tvNode.Parent & "]", "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & "", adOpenStatic, adLockReadOnly
      Set DataGrid1.DataSource = rsTView
      DataGrid1.Visible = True
      DataGrid2.Visible = False
      fraGrid.Visible = True
      fraEmng.Visible = False
      fraSearch.Visible = False
      fraInfo.Visible = False
      rtBox.Visible = False
      cmdRtbox.Visible = False
    Else
    End If

Exit Sub
eh_tvw:
  MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub txtQuery_Change()
' I think that there are some other ways to avoid execution of Select statement
' within connection , but this just shows to users, not to execute DML data modelling language
' if they choose to execute some update , delete or insert
On Error GoTo eh_query
   If optDML.Value = True Then
       If LTrim(txtQuery.Text) = "Select" Then
            cmdExecute.Enabled = False
            MsgBox "For Select Statements , Please use Standard SQL Query Option !", vbInformation, "Query Builder"
            txtQuery.Text = ""
            optDML.Value = False
            optQuery.Value = True
       ElseIf LTrim(txtQuery.Text) = "select" Then
            cmdExecute.Enabled = False
            MsgBox "For Select Statements , Please use Standard SQL Query Option !", vbInformation, "Query Builder"
            txtQuery.Text = ""
            optDML.Value = False
            optQuery.Value = True
       ElseIf LTrim(txtQuery.Text) = "SELECT" Then
            cmdExecute.Enabled = False
            MsgBox "For Select Statements , Please use Standard SQL Query Option !", vbInformation, "Query Builder"
            txtQuery.Text = ""
            optDML.Value = False
            optQuery.Value = True
        Else
            cmdExecute.Enabled = True
        End If
   Else
   End If

Exit Sub

eh_query:
  MsgBox Err.Description, vbCritical, "Query Builder"
End Sub

Private Sub txtSearch_Change()
    On Error GoTo eh_search
       
       lstResult.Clear
       Key = Trim(UCase(txtSearch.Text))
       strFind = "Select * from " & "[" & cmbTables.Text & "]" & " Where " & "[" & cmbColumns.Text & "]" & " LIKE '" & Key & "%'"
       rsCol.Open strFind, "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & user & ";Password=" & Pass & ";Initial Catalog=" & dbName & ";Data Source=" & srvName & "", adOpenStatic, adLockReadOnly

    Do While Not rsCol.EOF
       lstResult.AddItem rsCol.Fields(cmbColumns.Text)
       rsCol.MoveNext
    Loop
    
    If rsCol.RecordCount = 0 Then
        lblRecords.Caption = "Records found :" & rsCol.RecordCount
        MsgBox " No such Record !"
    Else
        lblRecords.Caption = "Records found :" & rsCol.RecordCount
    End If
    rsCol.Close
  
  Exit Sub
eh_search:
  MsgBox Err.Description, vbCritical, "Error"
End Sub
