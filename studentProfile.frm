VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   Caption         =   "Student Profile"
   ClientHeight    =   8205
   ClientLeft      =   6180
   ClientTop       =   1095
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9165
   Begin VB.CommandButton findbtn 
      Caption         =   "FIND"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   27
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton lastbtn 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   26
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton firstbtn 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   25
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton previousbtn 
      Caption         =   "<< PREVIOUS"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   24
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT >>"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   23
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   22
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   21
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   20
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton addbtn 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   19
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   7320
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   6480
      Width           =   3615
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   7
      Text            =   "Select Semester"
      Top             =   5640
      Width           =   3615
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   6
      Text            =   "Select Course"
      Top             =   4800
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   5
      Text            =   "Select Department"
      Top             =   3960
      Width           =   3615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   84344833
      CurrentDate     =   43192
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label9 
      Caption         =   "Phone No :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Semester :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Course :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Dept :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "DOB :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Gender :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Roll No :"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub addbtn_Click()
    rs.AddNew
    clear
End Sub
Sub clear()
    Text1.Text = ""
    Text2.Text = ""
    DTPicker1.Value = "01/01/1950"
    Option1.Value = False
    Option2.Value = False
    Combo1.Text = "Select the Department"
    Combo2.Text = "Select the Course"
    Combo3.Text = "Select the Semester"
    Text3.Text = ""
    Text4.Text = ""
End Sub
Private Sub Combo1_Click()
    Combo2.clear
    If Combo1.Text = "ENTC" Then
        Combo2.AddItem "BSc"
        Combo2.AddItem "BSc(Hons.)"
    ElseIf Combo1.Text = "CS" Then
        Combo2.AddItem "BSc"
        Combo2.AddItem "M.C.A"
        Combo2.AddItem "B.C.A"
    ElseIf Combo1.Text = "Electrical" Then
        Combo2.AddItem "BSc"
        Combo2.AddItem "BSc(Elec)"
    ElseIf Combo1.Text = "Civil" Then
        Combo2.AddItem "BSc"
        Combo2.AddItem "BSc(Civil)"
    ElseIf Combo1.Text = "Mechanical" Then
        Combo2.AddItem "BSc"
        Combo2.AddItem "BSc(Mech.)"
    ElseIf Combo1.Text = "Material" Then
        Combo2.AddItem "BSc"
        Combo2.AddItem "BSc(Mat.)"
    ElseIf Combo1.Text = "Chemical" Then
        Combo2.AddItem "BSc"
        Combo2.AddItem "BSc(Chem.)"
        
    End If
        
End Sub
Sub refreshD()
    rs.Close
    rs.Open "Select * from ProfileTBL", con, adOpenStatic, adLockPessimistic
    If rs.EOF Then
        MsgBox "No Records Found", vbInformation, "Message"
    Else
        rs.MoveNext
        If rs.EOF Then
            rs.MoveFirst
            display
        Else
            display
        End If
    End If
End Sub

Private Sub deletebtn_Click()
    Dim x As String
    x = MsgBox("Do you really want to delete thid record ?", vbCritical + vbYesNo, "Delete")
    If x = vbYes Then
        rs.Delete adAffectCurrent
        MsgBox "Record Deleted Successfully!", vbInformation, "Message"
        rs.Update
        refreshD
    Else
    End If
    
End Sub
Sub reload()
    rs.Close
    rs.Open "Select * from ProfileTBL", con, adOpenDynamic, adLockPessimistic
End Sub
Private Sub findbtn_Click()
    rs.Close
    rs.Open "Select * from ProfileTBL where RollNo = '" + Text1.Text + "'", con, adOpenDynamic, adLockPessimistic
    If rs.EOF Then
        MsgBox "No Record Found !!", vbInformation, "Message"
        refreshD
    Else
        display
    End If
End Sub

Private Sub firstbtn_Click()
    rs.MoveFirst
    display
End Sub

Private Sub Form_Load()

    con.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source =ProfileDB.mdb;Persist Security Info = False"
    rs.Open "Select * from ProfileTBL", con, adOpenDynamic, adLockPessimistic
    
    
    Combo1.AddItem "ENTC"
    Combo1.AddItem "CS"
    Combo1.AddItem "Electrical"
    Combo1.AddItem "Civil"
    Combo1.AddItem "Mechanical"
    Combo1.AddItem "Material"
    Combo1.AddItem "Chemical"
    
    Combo3.AddItem "Semester I"
    Combo3.AddItem "Semester II"
    Combo3.AddItem "Semester III"
    Combo3.AddItem "Semester IV"
    Combo3.AddItem "Semester V"
    Combo3.AddItem "Semester VI"
    Combo3.AddItem "Semester VII"
    Combo3.AddItem "Semester VIII"
    display
    
End Sub
Sub display()
    Text1.Text = rs!RollNo
    Text2.Text = rs!Name
    DTPicker1.Value = rs!DOB
    If rs!Gender = "Male" Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    Combo1.Text = rs!Dept
    Combo2.Text = rs!Course
    Combo3.Text = rs!Semester
    Text3.Text = rs!Address
    Text4.Text = rs!Phone
End Sub

Private Sub lastbtn_Click()
    rs.MoveLast
    display
End Sub

Private Sub nextbtn_Click()
    rs.MoveNext
    If rs.EOF Then
        rs.MoveFirst
        display
    Else
        display
    End If
End Sub

Private Sub previousbtn_Click()
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveLast
        display
    Else
        display
    End If
End Sub

Private Sub savebtn_Click()
    rs.Fields("RollNo").Value = Text1.Text
    rs.Fields("Name").Value = Text2.Text
    rs.Fields("DOB").Value = DTPicker1.Value
    If Option1.Value = True Then
        rs.Fields("Gender") = Option1.Caption
    Else
        rs.Fields("Gender") = Option2.Caption
    End If
    rs.Fields("Dept").Value = Combo1.Text
    rs.Fields("Course").Value = Combo2.Text
    rs.Fields("Semester").Value = Combo3.Text
    rs.Fields("Address").Value = Text3.Text
    rs.Fields("Phone").Value = Text4.Text
    MsgBox "Suceeded !!", vbInformation, "Message"
    rs.Update
    
End Sub

Private Sub uploadbtn_Click()
    Dim str As String
    CommonDialog1.ShowOpen
    CommonDialog1.Filter = "jpg|*jpg"
    str = CommonDialog1.FileName
    Picture1.Picture = LoadPicture(str)
End Sub

Private Sub updatebtn_Click()
    rs.Fields("RollNo").Value = Text1.Text
    rs.Fields("Name").Value = Text2.Text
    rs.Fields("DOB").Value = DTPicker1.Value
    If Option1.Value = True Then
        rs.Fields("Gender") = Option1.Caption
    Else
        rs.Fields("Gender") = Option2.Caption
    End If
    rs.Fields("Dept").Value = Combo1.Text
    rs.Fields("Course").Value = Combo2.Text
    rs.Fields("Semester").Value = Combo3.Text
    rs.Fields("Address").Value = Text3.Text
    rs.Fields("Phone").Value = Text4.Text
    MsgBox "Record Updated Successfully!", vbInformation, "Message"
    rs.Update
End Sub
