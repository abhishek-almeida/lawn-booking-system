VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form main_form 
   Caption         =   "Lawn Booking"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabbed_menu 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   16536
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Home | Overview"
      TabPicture(0)   =   "main_form.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frame_details"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frame_menu"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frame_reports"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Manage Bookings"
      TabPicture(1)   =   "main_form.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frame_manage"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frame_cost"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frame_event"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frame_client"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "History"
      TabPicture(2)   =   "main_form.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btn_modify"
      Tab(2).Control(1)=   "btn_show_all"
      Tab(2).Control(2)=   "btn_show_range"
      Tab(2).Control(3)=   "show_till_date"
      Tab(2).Control(4)=   "show_from_date"
      Tab(2).Control(5)=   "btn_last_week"
      Tab(2).Control(6)=   "history"
      Tab(2).Control(7)=   "lbl_show_till"
      Tab(2).Control(8)=   "lbl_show_from"
      Tab(2).ControlCount=   9
      Begin VB.Frame frame_reports 
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74640
         TabIndex        =   66
         Top             =   5880
         Width           =   11175
         Begin VB.CommandButton btn_revenue_report 
            Caption         =   "Revenue Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   8040
            TabIndex        =   69
            Top             =   720
            Width           =   2535
         End
         Begin VB.CommandButton btn_booking_records 
            Caption         =   "Booking Records"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   4410
            TabIndex        =   68
            Top             =   720
            Width           =   2595
         End
         Begin VB.CommandButton btn_client_info 
            Caption         =   "Client Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   600
            TabIndex        =   67
            Top             =   720
            Width           =   2520
         End
      End
      Begin VB.CommandButton btn_modify 
         Caption         =   "Update / Delete"
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
         Left            =   -70567
         TabIndex        =   64
         Top             =   7440
         Width           =   3015
      End
      Begin VB.CommandButton btn_show_all 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -66000
         TabIndex        =   57
         Top             =   780
         Width           =   2520
      End
      Begin VB.CommandButton btn_show_range 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -71520
         TabIndex        =   56
         Top             =   780
         Width           =   2520
      End
      Begin MSComCtl2.DTPicker show_till_date 
         Height          =   375
         Left            =   -72960
         TabIndex        =   55
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   156303361
         CurrentDate     =   44339
      End
      Begin MSComCtl2.DTPicker show_from_date 
         Height          =   375
         Left            =   -74520
         TabIndex        =   54
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   156303361
         CurrentDate     =   44339
      End
      Begin VB.CommandButton btn_last_week 
         Caption         =   "This Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -68760
         TabIndex        =   53
         Top             =   780
         Width           =   2520
      End
      Begin MSComctlLib.ListView history 
         Height          =   5430
         Left            =   -74640
         TabIndex        =   43
         Top             =   1740
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9578
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3000
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PHONE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "EVENT"
            Object.Width           =   3000
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "FROM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "TILL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "BOOKED ON"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "TOTAL COST"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame frame_client 
         Caption         =   "Client"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   360
         TabIndex        =   32
         Top             =   540
         Width           =   4695
         Begin VB.TextBox txt_email 
            Height          =   375
            Left            =   1680
            TabIndex        =   37
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txt_phone 
            Height          =   375
            Left            =   1680
            TabIndex        =   36
            Top             =   2040
            Width           =   2430
         End
         Begin VB.TextBox txt_address 
            Height          =   495
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   1440
            Width           =   2430
         End
         Begin VB.TextBox txt_age 
            Height          =   375
            Left            =   1680
            TabIndex        =   34
            Top             =   960
            Width           =   510
         End
         Begin VB.TextBox txt_name 
            Height          =   375
            Left            =   1680
            TabIndex        =   33
            Top             =   480
            Width           =   2430
         End
         Begin VB.Label lbl_booking_id 
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lbl_email 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   42
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label lbl_phone 
            Caption         =   "Phone"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   41
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lbl_address 
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   40
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lbl_age 
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   39
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lbl_name 
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   38
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame frame_menu 
         Caption         =   "Main Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74640
         TabIndex        =   28
         Top             =   720
         Width           =   4215
         Begin VB.CommandButton btn_history 
            Caption         =   "History"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   840
            TabIndex        =   31
            Top             =   2040
            Width           =   2535
         End
         Begin VB.CommandButton btn_manage 
            Caption         =   "Manage Bookings"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   840
            TabIndex        =   30
            Top             =   480
            Width           =   2535
         End
         Begin VB.CommandButton btn_about 
            Caption         =   "About Lawn Booking"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   840
            TabIndex        =   29
            Top             =   3480
            Width           =   2535
         End
      End
      Begin VB.Frame frame_details 
         Caption         =   "Overview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -70080
         TabIndex        =   27
         Top             =   720
         Width           =   6615
         Begin MSComctlLib.ListView daily_summary 
            Height          =   4095
            Left            =   360
            TabIndex        =   65
            Top             =   480
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   7223
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "NAME"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "EVENT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "FROM"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "TILL"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame frame_event 
         Caption         =   "Event"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   360
         TabIndex        =   13
         Top             =   4140
         Width           =   4695
         Begin VB.CommandButton btn_calculate 
            Caption         =   "Calculate Cost"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   1200
            TabIndex        =   58
            Top             =   3120
            Width           =   2295
         End
         Begin VB.ComboBox combo_type 
            Height          =   315
            ItemData        =   "main_form.frx":0054
            Left            =   1680
            List            =   "main_form.frx":006A
            Sorted          =   -1  'True
            TabIndex        =   22
            Top             =   480
            Width           =   2430
         End
         Begin VB.CheckBox check_decoration 
            Caption         =   "Decoration"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   7
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   17
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CheckBox check_catering 
            Caption         =   "Catering"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   7
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   16
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox check_dj 
            Caption         =   "DJ"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   7
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   15
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txt_count 
            Height          =   375
            Left            =   3000
            TabIndex        =   14
            Top             =   2520
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker time_till 
            Height          =   375
            Left            =   3000
            TabIndex        =   18
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "hh:mm tt"
            Format          =   125894659
            UpDown          =   -1  'True
            CurrentDate     =   44333
         End
         Begin MSComCtl2.DTPicker time_from 
            Height          =   375
            Left            =   3000
            TabIndex        =   19
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "hh:mm tt"
            Format          =   125894659
            UpDown          =   -1  'True
            CurrentDate     =   44333
         End
         Begin MSComCtl2.DTPicker date_till 
            Height          =   375
            Left            =   1680
            TabIndex        =   20
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   125894657
            CurrentDate     =   44336
         End
         Begin MSComCtl2.DTPicker date_from 
            Height          =   375
            Left            =   1680
            TabIndex        =   21
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   125894657
            CurrentDate     =   44336
         End
         Begin VB.Label lbl_type 
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lbl_from 
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   25
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lbl_till 
            Caption         =   "Till"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   24
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lbl_count 
            Caption         =   "Estimated Guests"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   23
            Top             =   2520
            Width           =   2415
         End
      End
      Begin VB.Frame frame_cost 
         Caption         =   "Cost Reciept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   5400
         TabIndex        =   6
         Top             =   540
         Width           =   6135
         Begin VB.Label lbl_booked_on 
            Alignment       =   2  'Center
            Caption         =   "Booked on: N/A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   59
            Top             =   3840
            Width           =   4095
         End
         Begin VB.Label lbl_decoration_rate 
            Alignment       =   2  'Center
            Caption         =   "- - -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   52
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lbl_dj_rate 
            Alignment       =   2  'Center
            Caption         =   "@5,000 Rs / hour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   51
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label lbl_catering_rate 
            Alignment       =   2  'Center
            Caption         =   "@800 Rs / plate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   50
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lbl_booking_rate 
            Alignment       =   2  'Center
            Caption         =   "@10,000 Rs / hour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   49
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lbl_total_charge 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   48
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label lbl_dj_charge 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   47
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label lbl_catering_charge 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   46
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lbl_decoration_charge 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   45
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lbl_booking_charge 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   44
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lbl_booking 
            Caption         =   "Booking Charge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   12
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lbl_decoration 
            Caption         =   "Decoration"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lbl_catering 
            Caption         =   "Catering"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lbl_dj 
            Caption         =   "DJ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label lbl_total 
            Caption         =   "Total Cost"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label lbl_event 
            Alignment       =   2  'Center
            Caption         =   "Event for x hours"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   7
            Top             =   480
            Width           =   4935
         End
         Begin VB.Line Line1 
            X1              =   600
            X2              =   5520
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line2 
            X1              =   600
            X2              =   5640
            Y1              =   3000
            Y2              =   3000
         End
      End
      Begin VB.Frame frame_manage 
         Caption         =   "Booking Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   5400
         TabIndex        =   1
         Top             =   5460
         Width           =   6135
         Begin VB.CommandButton btn_new 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   600
            TabIndex        =   60
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton btn_create 
            Caption         =   "Create"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2280
            TabIndex        =   5
            Top             =   600
            Width           =   1320
         End
         Begin VB.CommandButton btn_update 
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   600
            TabIndex        =   4
            Top             =   1680
            Width           =   1320
         End
         Begin VB.CommandButton btn_delete 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2280
            TabIndex        =   3
            Top             =   1680
            Width           =   1320
         End
         Begin VB.CommandButton btn_print 
            Caption         =   "Print Reciept"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   3960
            TabIndex        =   2
            Top             =   1080
            Width           =   1800
         End
      End
      Begin VB.Label lbl_show_till 
         Caption         =   "Till:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72960
         TabIndex        =   63
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lbl_show_from 
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   62
         Top             =   540
         Width           =   1095
      End
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cost_calculated As Boolean

Public history_sql As String
Public client_id_sql As String
Public event_from As Date
Public event_till As Date
Public estimated_guests As Integer
Public total_booking_charge As Long
Public event_duration As Single
Public decoration_charge As Long
Public catering_charge As Long
Public dj_charge As Long
Public total_cost As Long
Public booked_on As Date
Public show_booking_date As Boolean

Private Sub btn_about_Click()
    ans = MsgBox("Lawn Booking System ver1.0" & vbNewLine & vbNewLine & "Developed By -" & vbNewLine & "Abhishek Almeida", vbInformation, "Lawn Booking")
End Sub

Private Sub btn_booking_records_Click()
    sql_str = "select * from client_info c, booking_info b where c.client_id = b.client_id"
    Set rs = cn.Execute(sql_str)
    If Not rs.EOF Then
        Set booking_records.DataSource = rs
        booking_records.Show
    End If
End Sub

Private Sub btn_calculate_Click()
    On Error GoTo resolve_error
    
    Const booking_charge_per_hr As Integer = 10000
    Const per_plate_charge As Integer = 800
    Const dj_charge_per_hr As Integer = 5000
        
    show_booking_date = False
        
    Dim from As String
    Dim till As String
    
    from = Format(date_from.Value, "yyyy-MM-dd") & " " & Format(time_from.Value, "hh:mm:ss")
    till = Format(date_till.Value, "yyyy-MM-dd") & " " & Format(time_till.Value, "hh:mm:ss")
    
    event_from = CDate(from)
    event_till = CDate(till)
    
    event_duration = (DateDiff("n", event_from, event_till)) / 60
    If Len(combo_type.Text) < 3 Then
        Err.Raise _
            Number:=100, _
            Description:="Please enter proper event type."
    End If
    
    If event_duration <= 0 Then
        Err.Raise _
            Number:=101, _
            Description:="Event duration invalid. Please enter proper date and time."
    End If
   
    
    If check_decoration.Value = vbChecked Then
        Select Case LCase(combo_type.Text)
            Case "birthday party":
                decoration_charge = 20000
            Case "corporate party":
                decoration_charge = 80000
            Case "engagement":
                decoration_charge = 150000
            Case "farewell party":
                decoration_charge = 50000
            Case "festival celebration":
                decoration_charge = 30000
            Case "wedding reception":
                decoration_charge = 200000
            Case Else:
enter_decoration_charge:
                decoration_charge = InputBox("Decoration Charge", "What the decoration charge should be?")
                If decoration_charge <= 0 Then
                    Err.Raise _
                        Number:=102, _
                        Description:="Please enter proper decoration charge."
                End If
        End Select
    End If
    
    estimated_guests = Val(txt_count.Text)
        If estimated_guests <= 0 Then
            Err.Raise _
                Number:=103, _
                Description:="Please enter the number of estimated guests."
        End If
    
    If check_catering.Value = vbChecked Then
        catering_charge = per_plate_charge * CLng(estimated_guests)
    End If
    
    If check_dj.Value = vbChecked Then
        dj_charge = dj_charge_per_hr * event_duration
    End If

    total_booking_charge = booking_charge_per_hr * event_duration
    
    total_cost = total_booking_charge + decoration_charge + catering_charge + dj_charge
        
    cost_calculated = True
    generate_receipt
    
resolve_error:
    Select Case Err.Number
        Case 100
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            Exit Sub
        Case 101
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            Exit Sub
        Case 102
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            Resume enter_decoration_charge
        Case 103
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            Exit Sub
    End Select
End Sub

Sub generate_receipt()
    ' grab info from event details and show them in reciept section
    lbl_event.Caption = combo_type.Text & " for " & event_duration & " hours"
    lbl_booking_charge.Caption = total_booking_charge
    lbl_decoration_charge.Caption = decoration_charge
    lbl_catering_charge.Caption = catering_charge
    lbl_dj_charge.Caption = dj_charge
    lbl_total_charge.Caption = total_cost
    
    ' if record exists in backend, then show booking date else display n/a
    If show_booking_date = True Then
        lbl_booked_on.Caption = "Booked on: " & booked_on
    Else
        lbl_booked_on.Caption = "Booked on: N/A"
    End If
End Sub

Private Sub btn_client_info_Click()
    ' sql query to select all client info
    sql_str = "select * from client_info"
    Set rs = cn.Execute(sql_str)
    ' if data exists
    If Not rs.EOF Then
        Set client_info.DataSource = rs
        client_info.Show
    End If
End Sub

Private Sub btn_create_Click()
    
    Dim rs As adodb.Recordset
    Dim create_sql As String

    ' execute validation function to check if entered data is in correct format or not
    If validate_client_details = False Then
        Exit Sub
    End If
    
    ' sql query to insert client info
    client_insert_sql = "insert into client_info (name, age, address, phone, email) values " _
                        & "('" & txt_name.Text & "'," _
                        & txt_age.Text & "," _
                        & "'" & txt_address.Text & "'," _
                        & txt_phone.Text & "," _
                        & "'" & txt_email.Text & "')"
                        
    cn.Execute (client_insert_sql) ' execute client insert query
    
    ' grab client_id from client_info table (to insert it into booking_info table)
    client_id_sql = "select client_id from client_info where phone = " & txt_phone.Text
    
    ' sql query to insert booking info
    booking_insert_sql = "insert into booking_info (client_id, event_type, from_date, till_date, duration, booking_charge, decoration_charge, catering_charge, dj_charge, est_guests, booked_on, total_cost) values " _
                        & "(" & get_client_id(client_id_sql) & "," _
                        & "'" & combo_type.Text & "'," _
                        & "'" & Format(event_from, "yyyy-MM-dd hh:mm:ss") & "'," _
                        & "'" & Format(event_till, "yyyy-MM-dd hh:mm:ss") & "'," _
                        & event_duration & "," _
                        & total_booking_charge & "," _
                        & decoration_charge & "," _
                        & catering_charge & "," _
                        & dj_charge & "," _
                        & estimated_guests & "," _
                        & "Now()," _
                        & total_cost & ")" _

    cn.Execute (booking_insert_sql) ' execute booking insert query
    
    ' display confirmation message
    ans = MsgBox("Booking Created!", vbInformation, "Lawn Booking")
    
    btn_create.Enabled = False ' disable create button to avoid creating the same record again
    
    
    refresh_app ' custom function in general

End Sub

Private Sub btn_delete_Click()
    
    ' check if data imported from backend
    If check_booking_id = False Then
        Exit Sub
    End If

    ' check if all details are in proper format
    If validate_client_details = False Then
        Exit Sub
    End If
    
    ' confirmation prompt
    ans = MsgBox("Are you sure you want to delete this booking?", vbExclamation + vbYesNo, "Confirm")
    If ans = vbNo Then
        Exit Sub
    ElseIf ans = vbYes Then
        ' delete record sql
        delete_booking_sql = "delete from booking_info where booking_id = " & lbl_booking_id.Caption
        
        cn.Execute (delete_booking_sql)
        ans = MsgBox("Booking deleted!", vbInformation, "Lawn Booking")
        refresh_app
    End If
    
    
End Sub

Private Sub btn_history_Click()
    tabbed_menu.Tab = 2 ' switch to tab 2
End Sub

Private Sub btn_last_week_Click()
    Dim date_sub As Date
    date_sub = DateAdd("d", -Weekday(Now), Now) ' set date as last week day
    history_sql = "select * " _
        & "from client_info c, booking_info b " _
        & "where c.client_id = b.client_id " _
        & "and from_date >= '" & Format(date_sub, "yyyy-MM-dd") & " 00:00:00'" _
        & " and till_date <= now()"
    generate_history (history_sql)
End Sub

Private Sub btn_manage_Click()
    tabbed_menu.Tab = 1 ' switch to tab 1
End Sub

Private Sub btn_modify_Click()
    tabbed_menu.Tab = 1
End Sub

Private Sub btn_new_Click()
    btn_create.Enabled = True ' can create new records
    clear_all  ' execute function that clears all fields
End Sub

Private Sub btn_print_Click()
    ' check if record imported from backend
    If check_booking_id = False Then
        Exit Sub
    End If
    ' select all user details
    sql_str = "select * from client_info c, booking_info b where c.client_id = b.client_id" _
            & " and booking_id = " & lbl_booking_id.Caption
    Set rs = cn.Execute(sql_str)
    ' if record exists
    If Not rs.EOF Then
        Set reciept.DataSource = rs
        reciept.Show
    End If

End Sub

Private Sub btn_revenue_report_Click()
    sql_str = "select monthname(booked_on) as month, year(booked_on) as year, " _
                & "sum(total_cost) as earning from booking_info " _
                & "group by year, month desc " _
                & "union " _
                & "select 'TOTAL', '=', sum(total_cost) as total_earning from booking_info"
    Set rs = cn.Execute(sql_str)
    If Not rs.EOF Then
        Set revenue_report.DataSource = rs
        revenue_report.Show
    End If
End Sub

Private Sub btn_show_all_Click()
    refresh_app
End Sub

Private Sub btn_show_range_Click()
    ' show records between a certain range
    history_sql = "select * " _
       & "from client_info c, booking_info b " _
       & "where c.client_id = b.client_id " _
       & "and from_date >= '" & Format(show_from_date, "yyyy-MM-dd") & " 00:00:00' " _
       & "and till_date <= '" & Format(show_till_date, "yyyy-MM-dd") & " 00:00:00'"
    generate_history (history_sql)
End Sub

Private Sub btn_update_Click()

    If check_booking_id = False Then
        Exit Sub
    End If

    If validate_client_details = False Then
        Exit Sub
    End If
    
    ans = MsgBox("Are you sure you want to update this booking?", vbExclamation + vbYesNo, "Confirm")
    If ans = vbNo Then
        Exit Sub
    ElseIf ans = vbYes Then
        client_id_sql = "select client_id from booking_info where booking_id = " & lbl_booking_id.Caption
        
        update_client_sql = "update client_info " _
                    & "set name = '" & txt_name.Text & "'," _
                    & "age = " & txt_age.Text & "," _
                    & "address = '" & txt_address.Text & "'," _
                    & "phone = " & txt_phone.Text & "," _
                    & "email = '" & txt_email.Text & "'" _
                    & " where client_id = " & get_client_id(client_id_sql)
                    
        update_booking_sql = "update booking_info " _
                            & "set event_type = '" & combo_type.Text & "'," _
                            & "from_date = '" & Format(event_from, "yyyy-MM-dd hh:mm:ss") & "'," _
                            & "till_date = '" & Format(event_till, "yyyy-MM-dd hh:mm:ss") & "'," _
                            & "duration = " & event_duration & "," _
                            & "booking_charge = " & total_booking_charge & "," _
                            & "decoration_charge = " & decoration_charge & "," _
                            & "catering_charge = " & catering_charge & "," _
                            & "dj_charge = " & dj_charge & "," _
                            & "est_guests = " & estimated_guests & "," _
                            & "booked_on = now()," _
                            & "total_cost = " & total_cost _
                            & " where booking_id = " & lbl_booking_id.Caption
        cn.Execute (update_client_sql)
        cn.Execute (update_booking_sql)
        ans = MsgBox("Booking updated!", vbInformation, "Lawn Booking")
        refresh_app
    End If
    

End Sub

Private Sub Form_Load()
    tabbed_menu.Tab = 0
    date_from.Value = Now
    time_from.Value = Now
    date_till.Value = Now
    time_till.Value = Now
    refresh_app
End Sub

Sub generate_history(ByVal history_sql As String)
    Set rs = cn.Execute(history_sql)
    history.ListItems.Clear ' clear all list items
    Do While Not rs.EOF ' while records exist set different data fields
    Set Item = history.ListItems.Add(Text:=rs!booking_id)
    Item.SubItems(1) = rs!Name
    Item.SubItems(2) = rs!phone
    Item.SubItems(3) = rs!event_type
    Item.SubItems(4) = Format(rs!from_date, "dd/MM/yy, hh:mm")
    Item.SubItems(5) = Format(rs!till_date, "dd/MM/yy, hh:mm")
    Item.SubItems(6) = Format(rs!booked_on, "dd/MM/yy, hh:mm")
    Item.SubItems(7) = rs!total_cost
    rs.MoveNext ' select next record in record set object
    Loop
    Set rs = Nothing
End Sub

Private Sub history_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' on clicking any history item, set data field values to current item data
    With Item
        select_sql = "select * " _
                    & "from client_info c, booking_info b " _
                    & "where booking_id = " & Item _
                    & " and c.client_id = b.client_id"
        Set rs = cn.Execute(select_sql)
        lbl_booking_id.Caption = rs!booking_id
        txt_name.Text = rs!Name
        txt_age.Text = rs!age
        txt_address.Text = rs!address
        txt_phone.Text = rs!phone
        txt_email.Text = rs!email
        
        combo_type.Text = rs!event_type
        date_from.Value = rs!from_date
        time_from.Value = rs!from_date
        date_till.Value = rs!till_date
        time_till.Value = rs!till_date
        If rs!decoration_charge > 0 Then
            check_decoration.Value = vbChecked
        Else
            check_decoration.Value = vbUnchecked
        End If
        If rs!catering_charge > 0 Then
            check_catering.Value = vbChecked
        Else
            check_catering.Value = vbUnchecked
        End If
        If rs!dj_charge > 0 Then
            check_dj.Value = vbChecked
        Else
            check_dj.Value = vbUnchecked
        End If
        txt_count = rs!est_guests
        total_booking_charge = rs!booking_charge
        decoration_charge = rs!decoration_charge
        catering_charge = rs!catering_charge
        dj_charge = rs!dj_charge
        total_cost = rs!total_cost
        event_duration = rs!duration
        booked_on = rs!booked_on
        show_booking_date = True
        generate_receipt
        btn_create.Enabled = False ' disable button to avoid re-creation of record
    End With
End Sub

Sub clear_all()
    cost_calculated = False
    txt_name.Text = ""
    txt_age.Text = ""
    txt_address.Text = ""
    txt_phone.Text = ""
    txt_email.Text = ""
    combo_type.Text = ""
    date_from.Value = Now
    time_from.Value = Now
    date_till.Value = Now
    time_till.Value = Now
    check_decoration.Value = vbUnchecked
    check_catering.Value = vbUnchecked
    check_dj.Value = vbUnchecked
    txt_count.Text = ""
    lbl_event.Caption = "Event for x hours"
    lbl_booking_charge.Caption = ""
    lbl_decoration_charge.Caption = ""
    lbl_catering_charge.Caption = ""
    lbl_dj_charge.Caption = ""
    lbl_total_charge.Caption = ""
    lbl_booked_on.Caption = "Booked on: N/A"
    lbl_booking_id.Caption = ""
End Sub

Function validate_client_details() As Boolean
    On Error GoTo resolve_error
    If cost_calculated = False Then
            Err.Raise _
                Number:=110, _
                Description:="Please calculate cost first."
        End If
    
    If Len(txt_name.Text) = 0 _
    Or Len(txt_age.Text) = 0 _
    Or Len(txt_address.Text) = 0 _
    Or Len(txt_phone.Text) = 0 _
    Or Len(txt_email.Text) = 0 Then
        Err.Raise _
            Number:=111, _
            Description:="Incomplete details. Please enter all of the client details."
    End If
    
    Dim age_regex As RegExp
    Set age_regex = New RegExp
    age_regex.Pattern = "[0-9]{2,3}"
    
    If txt_age.Text < 18 Or Not age_regex.Test(txt_age.Text) Then
        Err.Raise _
            Number:=112, _
            Description:="Age is less than 18. Please enter proper age."
    End If
    
    Dim phone_regex As RegExp
    Set phone_regex = New RegExp
    phone_regex.Pattern = "[0-9]{10}"
    
    If Len(txt_phone.Text) <> 10 Or Not phone_regex.Test(txt_phone.Text) Then
        Err.Raise _
            Number:=113, _
            Description:="Invalid phone number. Please enter proper phone."
        txt_phone.SetFocus
    End If
    
    Dim email_regex As RegExp
    Set email_regex = New RegExp
    email_regex.Pattern = "[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+.[a-zA-Z0-9-.]+"
        
    If Not email_regex.Test(txt_email.Text) Then
        Err.Raise _
            Number:=114, _
            Description:="Invalid email address. Please enter proper email."
    End If
    
    validate_client_details = True
    
resolve_error:
    Select Case Err.Number
        Case 110
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            Exit Function
        Case 111
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            Exit Function
        Case 112
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            txt_age.SetFocus
            Exit Function
        Case 113
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            txt_phone.SetFocus
            Exit Function
        Case 114
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            txt_email.SetFocus
            Exit Function
    End Select
End Function


Function get_client_id(ByVal str As String) As Integer
    Set rs = cn.Execute(str)
    Dim client_id As Integer
    client_id = rs!client_id
    Set rs = Nothing
    get_client_id = client_id
End Function


Sub refresh_app()
    clear_all
    history_sql = "select * " _
        & "from client_info c, booking_info b " _
        & "where c.client_id = b.client_id "
    generate_history (history_sql)
    generate_summary
End Sub

Function check_booking_id() As Boolean
    On Error GoTo resolve_error

    If lbl_booking_id.Caption = "" Then
        Err.Raise _
            Number:=101, _
            Description:="Please select a record from History"
    End If
    
    check_booking_id = True
    
resolve_error:
    Select Case Err.Number
        Case 101
            ans = MsgBox(Err.Description, vbExclamation, "Lawn Booking")
            tabbed_menu.Tab = 2
            Exit Function
    End Select
End Function

Sub generate_summary()
    summary_sql = "select * from client_info c, booking_info b " _
                & "where b.client_id = c.client_id " _
                & "and date(from_date) = curdate()"
    Set rs = cn.Execute(summary_sql)
    daily_summary.ListItems.Clear
    Do While Not rs.EOF
    Set Item = daily_summary.ListItems.Add(Text:=rs!booking_id)
    Item.SubItems(1) = rs!Name
    Item.SubItems(2) = rs!event_type
    Item.SubItems(3) = Format(rs!from_date, "dd/MM, hh:mm")
    Item.SubItems(4) = Format(rs!till_date, "dd/MM, hh:mm")
    rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Lawn Booking") = vbNo Then
        Cancel = True
    Exit Sub
    Else
        cn.Close
  End If
End Sub

