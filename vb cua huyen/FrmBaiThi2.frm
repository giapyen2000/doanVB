VERSION 5.00
Begin VB.Form FrmBaiThi 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chuong trinh thi trac nghiem"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15195
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnTime"
      Size            =   12.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmdnop 
      Caption         =   "N�p"
      Height          =   735
      Left            =   13800
      TabIndex        =   23
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdE 
      Caption         =   "&E"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   18
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame ketqua 
      BackColor       =   &H00800080&
      Caption         =   "ke�t qua� thi"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   3600
      TabIndex        =   12
      Top             =   1080
      Width           =   11655
      Begin VB.CommandButton thoat 
         Caption         =   "THOAT"
         Height          =   375
         Left            =   10320
         TabIndex        =   19
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KE�T QUA� BA�I THI"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�I�M S�:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   16
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000016&
         BorderWidth     =   2
         X1              =   120
         X2              =   11280
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Labsodiem 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   ".VnTime"
            Size            =   99.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2415
         Left            =   5160
         TabIndex        =   15
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Labsocausai 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S� C�U SAI:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   5400
         TabIndex        =   14
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Labsocaudung 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S� C�U ��NG :"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   480
         TabIndex        =   13
         Top             =   2160
         Width           =   3375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3720
      Top             =   7200
   End
   Begin VB.CommandButton CmdD 
      Caption         =   "&D"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   4
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton CmdC 
      Caption         =   "&C"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton CmdB 
      Caption         =   "&B"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton CmdA 
      Caption         =   "&A"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Txthiencauhoi 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   11655
   End
   Begin VB.Label Label6 
      Caption         =   "S� ph�t : 15"
      Height          =   375
      Left            =   13680
      TabIndex        =   22
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "S� c�u : 10"
      Height          =   375
      Left            =   13680
      TabIndex        =   21
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "M�y : 01"
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   20
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line5 
      X1              =   3480
      X2              =   3480
      Y1              =   -1440
      Y2              =   7200
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3120
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Labngaysinhsv 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ng�y sinh:"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Labhotensv 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H� v� t�n:"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Lablopsv 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L�p:"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Labmasv 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M� sinh vi�n:"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Image Imghinhanh 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   600
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2160
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MO�N THI: TRA�C NGHIE�M TO�NG H��P"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   9615
   End
   Begin VB.Label Labdanglamcau 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ang l�m: "
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   3255
   End
   Begin VB.Label Labthoigiancon 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Th�i gian c�n: "
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   7680
      Width           =   3255
   End
End
Attribute VB_Name = "FrmBaiThi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim P As Integer

Private Sub CmdA_Click()
If DAN(P) = 1 Then
SoCauDung = SoCauDung + 1
End If
Call hiencauhoi
End Sub

Private Sub CmdB_Click()
If DAN(P) = 2 Then
SoCauDung = SoCauDung + 1
End If
Call hiencauhoi

End Sub

Private Sub CmdC_Click()
If DAN(P) = 3 Then
SoCauDung = SoCauDung + 1
End If
Call hiencauhoi
End Sub

Private Sub CmdD_Click()
If DAN(P) = 4 Then
SoCauDung = SoCauDung + 1
End If
Call hiencauhoi
End Sub

Private Sub cmdE_Click()
If DAN(P) = 5 Then
SoCauDung = SoCauDung + 1
End If
Call hiencauhoi
End Sub

Private Sub Cmdnop_Click()

    kt = MsgBox("Ban da lam duoc " & P & "/ " & 10 & "  cau, ban co chac chan muon nop bai khong ? ", vbYesNo, "Nop bai")
        If kt = 6 Then
            Call ketthuc
End If

End Sub


Private Sub thoat_Click()
End
End Sub

Private Sub Form_Initialize()
Thoigiancon = 15
SoCauDung = 0
SoCauSai = 0
sodiem = 0

Me.ketqua.Visible = False
Call daocauhoi

Me.Labmasv.Caption = "m� sinh vi�n: " & StrMSV
Me.Labhotensv.Caption = "h� v� t�n: " & StrHOTEN
Me.Labngaysinhsv.Caption = "Ng�y sinh: " & StrNS
Me.Lablopsv.Caption = "L�p : " & StrLOP
Me.Labthoigiancon.Caption = "Th�i gian c�n" & Thoigiancon & " ph�t"
anh = App.Path & "\ANH\" & StrMSV & ".JPG"
                                If Dir(anh) <> "" Then
                                    Me.Imghinhanh.Picture = LoadPicture(anh)
                                Else
                                    no = App.Path & "\Anh\No.JPG"
                                    Imghinhanh.Picture = LoadPicture(no)
                                End If

P = 1
dung = ""
For F = 1 To 5
dung = dung & Chr(F + 64) & ") " & TBN(P, F) & vbCrLf
Next
Me.Txthiencauhoi.Text = HBN(P) & vbCrLf & vbCrLf & dung
Me.Labdanglamcau.Caption = "�ang l�m c�u :  " & P & "/10"
End Sub
Sub hiencauhoi()

P = P + 1
dung = ""
If P <= 10 Then
For F = 1 To 5
    dung = dung & Chr(F + 64) & ") " & TBN(P, F) & vbCrLf
Next
        Me.Txthiencauhoi.Text = HBN(P) & vbCrLf & vbCrLf & dung
        Me.Labdanglamcau.Caption = "�ang l�m c�u :" & P & "/10"
    Else
        Call ketthuc
     End If

End Sub


Private Sub Timer1_Timer()
Thoigiancon = Thoigiancon - 1
If Thoigiancon >= 0 Then
        Me.Labthoigiancon.Caption = "th��i gian co�n :" & Thoigiancon & " phu�t"
Else
Call ketthuc
End If
End Sub
Sub ketthuc()
Me.CmdA.Enabled = False
Me.CmdB.Enabled = False
Me.CmdC.Enabled = False
Me.CmdD.Enabled = False
Me.cmdE.Enabled = False
Me.Cmdnop.Enabled = False
Me.Timer1.Interval = 0
Me.ketqua.Visible = True
Me.Labsocaudung.Caption = "s� c�u ��ng : " & SoCauDung
Me.Labsocausai.Caption = "s� c�u l�m sai :     " & 10 - SoCauDung
sodiem = Round((SoCauDung * 10) / 10, 2)
Me.Labsodiem.Caption = sodiem
Call GhiDuLieu

End Sub
Sub GhiDuLieu()
Dim G As Integer
G = FreeFile
Open "C:\KetQuaThi.Txt" For Append As #G
Print #G, StrMSV & "-" & StrHOTEN & "-" & StrNS & "-" & StrLOP & "-" & sodiem
Close #G
End Sub
Sub CapNhatDuLieuVaoDaTa()
        Set Cnn = New ADODB.Connection
        Set Rs = New ADODB.Recordset
        Cnn.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source= " & App.Path & "\DATA\DATA.mdb"
        Sql = "SELECT DANHSACH.* From DANHSACH WHERE MSV='" & StrMSV & "';"
        Rs.Open (Sql), Cnn '
        If Rs.EOF = False Then
                K = Rs("MSV")
                 If Not K Then
                        sqldata = "UPDATE DANHSACH SET DANHSACH.DIEM = " & sodiem & " WHERE (((DANHSACH.MSV)='" & StrMSV & "'));"
                        Cnn.Execute sqldata
                End If
    End If
End Sub
