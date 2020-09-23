VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRegister 
   Caption         =   "Your Registration Form"
   ClientHeight    =   1650
   ClientLeft      =   4530
   ClientTop       =   2730
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   3975
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtRegCode 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblRegistering 
      Caption         =   "Registering..."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblRegCode 
      Caption         =   "Registration Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   865
      Width           =   1455
   End
   Begin VB.Label lblEmail 
      Caption         =   "Registration Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label lblName 
      Caption         =   "Registration Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
'    Program Coder  : PorkNBeans
'    Project        : Basic FSO & Inet Example
'
'    Description:     A Very Basic Example On How To Connect To Ftp Programs
'                     And How To Create A Registration Type Program, Also Shows
'                     Simple Use Of The FSO And How To Delete A File After It
'                     Has Been Created... This Can Easily Be Improved Upon By Making
'                     An Array So Multiple Files Can Be Uploaded Or Multiple Users
'                     Can Resgister. This Was Done To Be Very Basic As To Not Confuse
'                     The Beginner, Intermediate Users Will Already Know How To Implement
'                     Arrays & Make It So Multiple Users Can Register. I Feel This Gives
'                     The Beginner The Necessary Foot In The Door On FSO & The Inet Control
'                     Without Letting Them Skip How The Component Actually Works.
'                     I Commented As Much As Possible To Teach As Best I Can On What The
'                     Code Is Doing.
'
'                     More Experienced Users Can Easily Implement A Registration Program
'                     That Actually Requires A Valid Registration Number To Be Entered. I
'                     Left The Program Open to Endless Possibilities And To Spark Ideas For
'                     All Users.
'
'                     Please Vote If You Find This Worthwhile, I Think It Is Worth Some
'                     Votes.
'
'    Current Date   : January 14th, 2001
'-------------------------------------------------------------------------------------------

Option Explicit

Private Sub cmdCancel_Click()
        
        On Error GoTo cmdCancel_Click_Err

100     Me.Hide
        
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
           "in Project1.frmRegister.cmdCancel_Click " & _
           "at line " & Erl
        Resume Next
        
End Sub

Private Sub cmdRegister_Click()
        
        On Error GoTo cmdRegister_Click_Err
        
100     If txtName.Text = "" Or txtEmail.Text = "" Or txtRegCode.Text = "" Then

102         MsgBox "You must enter a Registration Name, Registration Email Address, & Registration Code!", vbOKOnly + vbInformation, "Registration Error"

104     Else

106         Open "C:\temp.dat" For Output As #1 'Open Temp.dat for editing
108         Print #1, frmMain.Caption & vbCrLf 'Add the caption of the main form to the temp file that way to make it easier to keep track of what program was registered
110         Print #1, "Registration Name: " & txtName.Text & vbCrLf 'Adds the name they entered in the name textbox to the file
112         Print #1, "Registration Email: " & txtEmail.Text & vbCrLf 'Adds the email address they entered in the email textbox to the file
114         Print #1, "Registration Code: " & txtRegCode.Text & vbCrLf 'Adds the registration code they entered in the email textbox to the file
116         Print #1, "Date & Time: " & Format(Now) 'Adds the date and time to the file to aid in tracking down when they registered
118         Close #1 'Close the file after we're done adding the information

120         lblRegistering.Visible = True
122         lblRegistering.Caption = "Registering ... "
124         Inet1.URL = "ftp://your.FTP.here"     'Connect to the ftp server
126         Inet1.UserName = "FTP Login Name" 'Connect with this login name
128         Inet1.Password = "FTP Login Password" 'Connect with this password
130         Inet1.Execute , "PUT C:\temp.dat regged.dat" 'Upload the file and rename it on the server to the name of registration for security reasons

        End If
        
        Exit Sub

cmdRegister_Click_Err:
        MsgBox Err.Description & vbCrLf & _
           "in Project1.frmRegister.cmdRegister_Click " & _
           "at line " & Erl
        Resume Next
        
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
        
        On Error GoTo Inet1_StateChanged_Err
    
        Dim FSO As New FileSystemObject 'Dim Our Variables And So They May be Called Upon When Needed

100     If State = icResponseCompleted Then 'If the transaction is completed

102         Inet1.Execute , "Close" 'Close the ftp connection

104         Set FSO = CreateObject("Scripting.FileSystemObject") 'Call And Set Our Variable To Create The Object
106         FSO.DeleteFile "C:\temp.dat" 'Use the FSO to delete the temp.dat file we created once the registration has completed successfully
108         Me.Hide 'Last but not least hide the form once everything else has been done

        End If
        
        Exit Sub

Inet1_StateChanged_Err:
        MsgBox Err.Description & vbCrLf & _
           "in Project1.frmRegister.Inet1_StateChanged " & _
           "at line " & Erl
        Resume Next
        
End Sub

Private Sub lblRegCode_Click()

End Sub
