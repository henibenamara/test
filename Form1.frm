VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update DataBase"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Server"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton btnDeconnect 
         BackColor       =   &H000000FF&
         Caption         =   "déconnecter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "vider la base de donnée !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CommandButton btnConnecter 
         BackColor       =   &H0000C000&
         Caption         =   "Connecter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtBD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "username :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Base de donnée :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Server (IP) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Dim rs As New Recordset


Private Sub btnConnecter_Click()

    Dim password As String
    Dim NomBasseDonne As String
    Dim Server As String
    Dim User As String
    
    
    Server = txtServer.Text
    
    NomBasseDonne = txtBD.Text
    
    password = txtPass.Text
    
    User = txtUser.Text
    
    
    
    con.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};" _
        & "Server=" & Server & ";" _
        & "Database=" & NomBasseDonne & ";" _
        & "User=" & User & ";" _
        & "Password= " & password & ";" _
        & "Option=3;"
    
    con.Open
    
    MsgBox "connection avec succee"
    btnConnecter.Visible = False
    Command1.Visible = True
    btnDeconnect.Visible = True
    
    
    
    txtBD.Enabled = False
    
    
    txtPass.Enabled = False
    txtServer.Enabled = False
    txtUser.Enabled = False
End Sub

Private Sub btnDeconnect_Click()
    con.Close
    
    txtBD.Enabled = True
    txtPass.Enabled = True
    txtServer.Enabled = True
    txtUser.Enabled = True
    
    txtBD.Text = ""
    txtPass.Text = ""
    txtServer.Text = ""
    txtUser.Text = ""
    
    btnConnecter.Visible = True
    Command1.Visible = False
    btnDeconnect.Visible = False
    
End Sub

Private Sub Command1_Click()

    sql = "DELETE FROM article"
    sql1 = "DELETE FROM banque"
    sql2 = "DELETE FROM befa"
    sql3 = "DELETE FROM belfa"
    sql4 = "DELETE FROM bl"
    sql5 = "DELETE FROM blf"
    sql6 = "DELETE FROM br"
    sql7 = "DELETE FROM brf"
    sql8 = "DELETE FROM cloturecaisse"
    sql9 = "DELETE FROM ebc"
    sql10 = "DELETE FROM ebcf"
    sql11 = "DELETE FROM ebe"
    sql12 = "DELETE FROM ebp"
    sql13 = "DELETE FROM eca"
    
    sql14 = "DELETE FROM eca1"
    sql15 = "DELETE FROM eca2"
    sql16 = "DELETE FROM eca3"
    sql17 = "DELETE FROM eca4"
    sql18 = "DELETE FROM eca5"
    sql19 = "DELETE FROM eca6"
    sql20 = "DELETE FROM eca7"
    sql21 = "DELETE FROM eca8"
    sql22 = "DELETE FROM eca9"
    sql23 = "DELETE FROM eca99"
    sql24 = "DELETE FROM edp"
    sql25 = "DELETE FROM efa"
    sql26 = "DELETE FROM emdv"
    
    sql27 = "DELETE FROM lca"
    sql28 = "DELETE FROM lca1"
    sql29 = "DELETE FROM lca2"
    sql30 = "DELETE FROM lca3"
    sql31 = "DELETE FROM lca4"
    sql32 = "DELETE FROM lca5"
    sql33 = "DELETE FROM lca6"
    sql34 = "DELETE FROM lca7"
    sql35 = "DELETE FROM lca8"
    sql36 = "DELETE FROM lca9"
    
    sql37 = "DELETE FROM lbr"
    sql38 = "DELETE FROM ldp"
    sql39 = "DELETE FROM lfa"
    sql40 = "DELETE FROM lignedepot"
    sql41 = "DELETE FROM lsd"
    
    sql42 = "DELETE FROM pfa"
    sql43 = "DELETE FROM mlfa"
    sql44 = "DELETE FROM mdc"
    sql45 = "DELETE FROM paie"
    sql46 = "DELETE FROM regl"
    sql47 = "DELETE FROM pefa"
    sql48 = "DELETE FROM reglsd"
    sql49 = "DELETE FROM lmdv"
    sql50 = "DELETE FROM esd"
    sql51 = "DELETE FROM lbl"
    sql52 = "DELETE FROM lbe"
    sql53 = "DELETE FROM rerf"
    sql54 = "DELETE FROM rrf"
    
    
    sql55 = "DELETE FROM client where  code != 9000 "
    
    codeDepot = "DL"
    CodeFamille = "P_F"
    sql56 = "DELETE FROM depot where code !='" & codeDepot & "' "
    
    sql57 = "DELETE FROM famille where code !='" & CodeFamille & "'  "
    sql58 = "DELETE FROM ldfp"
    sql59 = "DELETE FROM lbcf"
    
    
    
    con.Execute sql
    con.Execute sql1
    con.Execute sql2
    con.Execute sql3
    con.Execute sql4
    con.Execute sql5
    con.Execute sql6
    con.Execute sql7
    con.Execute sql8
    con.Execute sql9
    con.Execute sql10
    con.Execute sql11
    con.Execute sql12
    con.Execute sql13
    con.Execute sql14
    con.Execute sql15
    con.Execute sql16
    con.Execute sql17
    con.Execute sql18
    con.Execute sql19
    con.Execute sql20
    con.Execute sql21
    con.Execute sql22
    con.Execute sql23
    con.Execute sql24
    con.Execute sql25
    con.Execute sql26
    con.Execute sql27
    con.Execute sql28
    con.Execute sql29
    con.Execute sql30
    con.Execute sql31
    con.Execute sql32
    con.Execute sql33
    con.Execute sql34
    con.Execute sql35
    con.Execute sql36
    con.Execute sql37
    con.Execute sql38
    con.Execute sql39
    con.Execute sql40
    con.Execute sql41
    con.Execute sql42
    con.Execute sql43
    con.Execute sql44
    con.Execute sql45
    con.Execute sql46
    con.Execute sql47
    con.Execute sql48
    con.Execute sql49
    con.Execute sql50
    con.Execute sql51
    con.Execute sql52
    con.Execute sql53
    con.Execute sql54
    con.Execute sql55
    con.Execute sql56
    con.Execute sql57
    con.Execute sql58
    con.Execute sql59
    
    MsgBox "suppression des données avec succès"
    con.Close
    
    txtBD.Enabled = True
    txtPass.Enabled = True
    txtServer.Enabled = True
    txtUser.Enabled = True
    
    txtBD.Text = ""
    txtPass.Text = ""
    txtServer.Text = ""
    txtUser.Text = ""
    
    btnConnecter.Visible = True
    Command1.Visible = False
    btnDeconnect.Visible = False
    
    
    
End Sub

Private Sub Form_Load()
Command1.Visible = False
btnDeconnect.Visible = False

End Sub
