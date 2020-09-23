VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5670
      Left            =   180
      TabIndex        =   1
      Top             =   495
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10001
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim rs As ADODB.Recordset
    Dim rsLookup As ADODB.Recordset
    Dim cnn As ADODB.Connection
    Dim FileName As String
    
    'Modify the following line to point to your Northwind database!
    FileName = "C:\Program Files\Microsoft Visual Studio\VB98\Nwind.mdb"
    
    'Create a new connection
    Set cnn = New Connection
    
    'Necessary, else you see nothing in the grid.
    cnn.CursorLocation = adUseClient
    
    'Open the connection
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FileName & ";Persist Security Info=False"
    
    'Open the data recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * From Products", cnn, adOpenStatic, adLockOptimistic
    
    'Hook the datagrid to the data
    Set DataGrid1.DataSource = rs
    
    'Open the lookup recordset
    Set rsLookup = New ADODB.Recordset
    rsLookup.Open "SELECT * From Suppliers", cnn, adOpenStatic, adLockReadOnly
    
    'Hook the data combo control to both the
    'data recordset AND the lookup recordset!
    With DataCombo1
    
        '---> SET AT DESIGN TIME. <---
        '.Style = dbcDropdownList
        
        'The data recordset
        Set .DataSource = rs
        
        'The lookup recordset
        Set .RowSource = rsLookup
        
        'Field name in the data recordset.
        .BoundColumn = "SupplierID"
        
        'Field name in the lookup recordset.
        .DataField = "SupplierID"
        
         'Field name to display instead of the SupplierID
         .ListField = "CompanyName"
    End With
    
    'Final note - you must move OFF THE RECORD to see the change! :-/
End Sub

