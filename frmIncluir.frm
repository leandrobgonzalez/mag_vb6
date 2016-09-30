VERSION 5.00
Begin VB.Form frmIncluir 
   Caption         =   "Incluir"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cboSexo 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtIdade 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   600
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Idade:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
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
Attribute VB_Name = "frmIncluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
