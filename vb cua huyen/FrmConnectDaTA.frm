VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmConnectDaTA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "thong tin sinh vien"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "VNI-Times"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   5760
      Top             =   7800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdKetThuc 
      Cancel          =   -1  'True
      Caption         =   "&keát thuùc"
      Height          =   435
      Left            =   4320
      TabIndex        =   6
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton CmdBatDau 
      Caption         =   "&baét ñaàu thi"
      Height          =   435
      Left            =   1440
      TabIndex        =   5
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox TxtMaSo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "(nhaäp maõ sv roài aán enter ñeå xaùc nhaän laøm baøi)"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "maõ sinh vieân:"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label LabLop 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lôùp:"
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   4275
   End
   Begin VB.Label LabHoten 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hoï vaø teân:"
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   4275
   End
   Begin VB.Label LabNgaysinh 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngaøy sinh:"
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   4275
   End
   Begin VB.Image Imghinhanh 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRÖÔØNG ÑAÏI HOÏC KINH DOANH VAØ COÂNG NGHEÄ HAØ NOÄI"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "FrmConnectDaTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim Sql As String

Private Sub CmdBatDau_Click()
Rs ("HODEM") & " " & Rs("Ten") = StrHoVaTen
Rs("NGAYSINH") = StrNgaySinh
Rs("LOP") = StrLop
Me.Hide
FrmBaiThi2.Show
End Sub

Private Sub CmdKetThuc_Click()
End
End Sub
Private Sub Form_Load()
Me.LabHoten.Caption = ""
Me.LabNgaysinh.Caption = ""
Me.LabLop.Caption = ""
Imghinhanh.Picture = LoadPicture("")
Me.CmdBatDau.Enabled = False
End Sub
Private Sub TxtMaSo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Set Cnn = New ADODB.Connection
        Set Rs = New ADODB.Recordset
        Cnn.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\DATA\DATA.mdb"
        Sql = "SELECT DANHSACH.* From DANHSACH WHERE MSV='" & Me.TxtMaSo & "';"
        Rs.Open (Sql), Cnn
         If Rs.EOF = False Then
                K = Rs("MSV")
                 If Not K Then
                        Me.LabHoten.Caption = "Ho va ten : " & Rs("HODEM") & " " & Rs("TEN")
                        Me.LabNgaysinh.Caption = "Ngay sinh: " & Rs("NGAYSINH")
                        Me.LabLop.Caption = "Lop: " & Rs("LOP")
                        
                        Imghinhanh.Picture = LoadPicture("")
                        
                        Anh = App.Path & "\ANH\" & Rs("MSV") & ".JPG"
                                
                                If Dir(Anh) <> "" Then
                                    Me.Imghinhanh.Picture = LoadPicture(Anh)
                                Else
                                    no = App.Path & "\Anh\No.JPG"
                                    Imghinhanh.Picture = LoadPicture(no)
                                End If
                    
                    Me.CmdBatDau.Enabled = True
                    Me.CmdBatDau.SetFocus
                End If
        Else
        Me.LabHoten.Caption = ""
        Me.LabNgaysinh.Caption = ""
        Me.LabLop.Caption = ""
        Imghinhanh.Picture = LoadPicture("")
        Me.TxtMaSo.SetFocus
        Me.TxtMaSo.Text = ""
        Me.CmdBatDau.Enabled = False
        
        MsgBox "khoâng tìm thaáy sinh vieân!", vbOKOnly, "Thong bao"
     
     End If
End If
End Sub
