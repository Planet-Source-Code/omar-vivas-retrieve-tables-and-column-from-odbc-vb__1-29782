VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RetrieveSchema 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get Tables and Fields"
   ClientHeight    =   6210
   ClientLeft      =   2355
   ClientTop       =   570
   ClientWidth     =   5595
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   5595
   Begin VB.Frame Frame3 
      Caption         =   "ODBC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdNewODBC 
         Caption         =   "..."
         Height          =   300
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   1600
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Security"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   5175
         Begin VB.TextBox txtPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox txtUser 
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "UserName"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox CmbDSN 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "DSN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   5535
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   5535
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   5760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=VTIME"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "VTIME"
      OtherAttributes =   ""
      UserName        =   "INSUDB"
      Password        =   "INSUDB"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "RetrieveSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rsSchema As ADODB.Recordset
Dim fld As ADODB.Field
Dim rCriteria As Variant
Dim ShowODBC As Boolean

Private Sub CmbDSN_Click()
    Me.Height = 2500
End Sub

Private Sub Command1_Click()
    If txtUser.Text <> "" And txtPass.Text <> "" Then
        If ConectarODBC(CmbDSN.Text, txtUser.Text, txtPass.Text) Then
            Me.Height = 6600
            ShowODBC = True
        End If
    Else
        MsgBox "You must write the User and Password to in", vbOKOnly + vbExclamation
    End If
End Sub

Private Sub cmdNewODBC_Click()
    RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL odbccp32.cpl,,3", 1)
End Sub

Private Sub Form_Load()
    ShowODBC = False
    GetODBCList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ShowODBC Then
        rsSchema.Close
        Set rsSchema = Nothing
        cn.Close
        Set cn = Nothing
        Set fld = Nothing
    End If
End Sub

Private Sub List1_Click()
    List2.Clear
    
    rCriteria = Array(Empty, Empty, List1.Text, Empty)
    
    'Set rsSchema = cn.OpenSchema(adSchemaIndexes, rCriteria)
    Set rsSchema = cn.OpenSchema(adSchemaColumns, rCriteria)
    
    Debug.Print "Recordcount: " & rsSchema.RecordCount
    
    While Not rsSchema.EOF
    
        NameField = rsSchema!COLUMN_NAME & Space(30 - Len(rsSchema!COLUMN_NAME))
        
        Select Case rsSchema!DATA_TYPE
        Case 129
            TypeField = " STRING(" & Trim(Str$(rsSchema!CHARACTER_MAXIMUM_LENGTH)) & ")"
        Case 131
            If rsSchema!NUMERIC_SCALE > 0 Then
                TypeField = " NUMERIC(" & Trim(Str$(rsSchema!NUMERIC_PRECISION)) & "," & Trim(Str$(rsSchema!NUMERIC_SCALE)) & ")"
            Else
                TypeField = " NUMERIC(" & Trim(Str$(rsSchema!NUMERIC_PRECISION)) & ")"
            End If
        Case 135
            TypeField = " DATE"
        Case Else
            TypeField = rsSchema!DATA_TYPE
        End Select
        
        TypeField = TypeField & Space(20 - Len(TypeField))
        List2.AddItem NameField & vbTab & TypeField & vbTab & "Req: " & IIf(rsSchema!IS_NULLABLE, "True", "False")
'       For Each fld In rsSchema.Fields
'          Debug.Print fld.Name
'          Debug.Print fld.Value
'          If fld.Name = "COLUMN_NAME" Then
'            List2.AddItem fld.Value
'          End If
'          If fld.Name = "DATA_TYPE" Then
'            List2.AddItem fld.Value
'          End If
'
'         List2.AddItem rsSchema!COLUMN_NAME & " " & rsSchema!DATA_TYPE
'          Debug.Print "------------------------------------------------"
'       Next
       rsSchema.MoveNext
    Wend
End Sub

Function ConectarODBC(ByVal sDSN As String, ByVal sUser As String, ByVal sPass As String) As Boolean
    Set cn = New ADODB.Connection
    
    List1.Clear
    List2.Clear
    ConectarODBC = True
    
    On Error GoTo ErrorHandler
    With cn
       .Provider = "MSDASQL"   'default Provider=MSDASQL
       .CursorLocation = adUseServer
       .ConnectionString = "DSN=" & sDSN & ";UID=" & sUser & ";PWD=" & sPass & ";"
       .Open
    End With
    
    GoTo Seguir
ErrorHandler:
    MsgBox Err.Description
    ConectarODBC = False
    Exit Function
Seguir:
    
    'Pass in the table name to retrieve index info. The other
    'array parameters may be defined as follows:
    '    TABLE_CATALOG  (first parameter)
    '    TABLE_SCHEMA   (second)
    '    INDEX_NAME     (third)
    '    TYPE           (fourth)
    '    TABLE_NAME     (fifth, e.g. "employee")
    'rCriteria = Array(Empty, Empty, Empty, Empty, "employee")
    'rCriteria = Array(Empty, Empty, Empty, Empty, "ACCIDENT")
    'rCriteria = Array(Empty, "INSUDB", Empty, Empty, Empty)
    
    'rCriteria = Array(Empty, "INSUDB", Empty, "Table")
    rCriteria = Array(Empty, Empty, Empty, "Table")
    
    'rCriteria = Array(Empty, Empty, "ACCIDENT", Empty)
    
    'Set rsSchema = cn.OpenSchema(adSchemaIndexes, rCriteria)
    'Set rsSchema = cn.OpenSchema(adSchemaColumns, rCriteria)
    Set rsSchema = cn.OpenSchema(adSchemaTables, rCriteria)
    
    Debug.Print "Recordcount: " & rsSchema.RecordCount
    
    While Not rsSchema.EOF
          Debug.Print "==================================================="
    
       For Each fld In rsSchema.Fields
          'Debug.Print fld.Name
          'Debug.Print fld.Value
          If fld.Name = "TABLE_NAME" Then
            List1.AddItem fld.Value
          End If
          'Debug.Print "------------------------------------------------"
       Next
       rsSchema.MoveNext
    Wend

End Function

Sub GetODBCList()
    Dim oODBC As New ODBCTool.Dsn
    Dim aDSN() As String
        
        oODBC.GetDataSourceList aDSN
        For i = LBound(aDSN) To UBound(aDSN)
            CmbDSN.AddItem aDSN(i)
        Next
        CmbDSN.ListIndex = 0
    Set oODBC = Nothing
End Sub
