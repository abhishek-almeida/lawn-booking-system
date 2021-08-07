VERSION 5.00
Begin VB.Form main_form_window 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   833
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_revenue_report 
      Caption         =   "Revenue Report"
      Height          =   975
      Left            =   840
      TabIndex        =   5
      Top             =   4920
      Width           =   3495
   End
   Begin VB.CommandButton btn_history 
      Caption         =   "History"
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton btn_client_info 
      Caption         =   "Client Info"
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Frame current_bookings 
      Caption         =   "Frame2"
      Height          =   6135
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Frame main_menu 
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.CommandButton btn_manage_bookings 
         Caption         =   "Manage Bookings"
         Height          =   975
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Width           =   3495
      End
   End
End
Attribute VB_Name = "main_form_window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
