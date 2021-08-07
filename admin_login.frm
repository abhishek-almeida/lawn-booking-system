VERSION 5.00
Begin VB.Form admin_login 
   Caption         =   "Lawn Booking | Login"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_login 
      Caption         =   "Login"
      Height          =   495
      Left            =   1815
      TabIndex        =   4
      Top             =   1995
      Width           =   1575
   End
   Begin VB.TextBox txt_passwd 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2235
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1185
      Width           =   1935
   End
   Begin VB.TextBox txt_user 
      Height          =   495
      Left            =   2235
      TabIndex        =   2
      Top             =   465
      Width           =   1935
   End
   Begin VB.Label lbl_passwd 
      Caption         =   "Password"
      Height          =   255
      Left            =   1035
      TabIndex        =   1
      Top             =   1305
      Width           =   975
   End
   Begin VB.Label lbl_user 
      Caption         =   "User"
      Height          =   255
      Left            =   1035
      TabIndex        =   0
      Top             =   585
      Width           =   975
   End
End
Attribute VB_Name = "admin_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_login_Click()
    ' create a record set object to hold backend data
    Dim rs As adodb.Recordset
    ' sql query to get admin login details
    login_sql = "select passwd from admin_login where user='" + txt_user.Text + "'"
    ' execute login query
    Set rs = cn.Execute(login_sql)
    ' if data exists
    If Not rs.EOF Then
        If txt_passwd.Text = rs!passwd Then
            main_form.Show  ' show main form
            Unload Me   ' unload(hide) login form
        Else
            incorrect_login
        End If
    Else
        incorrect_login
    End If
End Sub


' When login credentials are incorrect, display a prompt, clear entered details and set focus
Public Sub incorrect_login()
    ans = MsgBox("Incorrect Login", vbExclamation, "Login")
    txt_user.Text = ""
    txt_passwd.Text = ""
    txt_user.SetFocus
End Sub
