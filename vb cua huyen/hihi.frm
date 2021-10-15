VERSION 5.00
Begin VB.Form FrmConnectDaTA 
   BackColor       =   &H00800080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "thong tin sinh vien"
   ClientHeight    =   6090
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
   ScaleHeight     =   6090
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7170
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   12
      Top             =   7590
      Width           =   1200
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
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lôùp :"
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ngaøy sinh :"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hoï vaø teân:"
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
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
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   1920
      TabIndex        =   4
      Top             =   4440
      Width           =   2835
   End
   Begin VB.Label LabHoten 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   2835
   End
   Begin VB.Label LabNgaysinh 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   1920
      TabIndex        =   2
      Top             =   3600
      Width           =   2835
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

  StrMSV = Rs("MSV")
        StrHOTEN = Rs("HODEM") & " " & Rs("TEN")
        StrNS = Rs("NGAYSINH")
        StrLOP = Rs("LOP")
        
        Me.Hide
        FrmBaiThi.Show

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
                        Me.LabHoten.Caption = Rs("HODEM") & " " & Rs("TEN")
                        Me.LabNgaysinh.Caption = Rs("NGAYSINH")
                        Me.LabLop.Caption = Rs("LOP")
                        
                        Imghinhanh.Picture = LoadPicture("")
                        
                        anh = App.Path & "\ANH\" & Rs("MSV") & ".JPG"
                                
                                If Dir(anh) <> "" Then
                                    Me.Imghinhanh.Picture = LoadPicture(anh)
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
