VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsulta 
   Caption         =   "Consulta"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dtgConsulta 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3625
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
   Begin VB.CommandButton btnFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboSexo 
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtIdade 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Sexo:"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Idade:"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFiltrar_Click()
    
    Dim sUrl As String
    Dim response As String
    Set sUrl = "api url/"
    
    If Len(txtNome.Text) > 0 Then
        Set sUrl = sUrl & "Nome=" & txtNome.Text & "&"
    End If
    
    If Len(txtIdade.Text) > 0 Then
        Set sUrl = sUrl & "Idade=" & CInt(txtIdade.Text) & "&"
    End If
    
    If selIndex > 0 Then
        Set sUrl = sUrl & "Sexo=" & cboSexo.List(cboSexo.ListIndex) & "&"
    End If
    
    Set sUrl = Left$(sUrl, Len(sUrl) - 1)
    
    Dim http As MSXML.xmlHttp
    Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")
    
    http.Open "GET", sUrl, False
    http.Send
    
    Set response = xmlHttp.responseText
    Set http = Nothing
    
End Sub
