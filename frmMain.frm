VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFBDBD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Embed - by eidos"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   9
      Top             =   3840
      Width           =   3345
   End
   Begin VB.TextBox txtFolderName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   1
      Text            =   "%THISDIRNAME%"
      Top             =   690
      Width           =   3345
   End
   Begin VB.TextBox txtDelimiter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   5
      Text            =   ","
      Top             =   2250
      Width           =   3345
   End
   Begin VB.PictureBox picPreview 
      Height          =   3405
      Left            =   240
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3345
      ScaleWidth      =   2685
      TabIndex        =   24
      Top             =   720
      Width           =   2745
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFBDBD&
      Caption         =   "Custom HTML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   22
      Top             =   4320
      Width           =   8295
      Begin VB.TextBox txtCustomHtml 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2145
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   780
         Width           =   7965
      End
      Begin VB.CommandButton cmdInsertObject 
         Caption         =   "Insert Object >>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6000
         TabIndex        =   10
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "You can also insert custom HTML tags such as hyperlinks, objects, scripts, etc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   150
         TabIndex        =   23
         Top             =   300
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      TabIndex        =   15
      Top             =   7710
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4380
      TabIndex        =   14
      Top             =   7710
      Width           =   1305
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5820
      TabIndex        =   12
      Top             =   7710
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7260
      TabIndex        =   13
      Top             =   7710
      Width           =   1305
   End
   Begin VB.TextBox txtBytesText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   8
      Text            =   "bytes"
      Top             =   3450
      Width           =   3345
   End
   Begin VB.TextBox txtTotalFileSizeText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   7
      Text            =   "Total File Size: "
      Top             =   3045
      Width           =   3345
   End
   Begin VB.TextBox txtSizeText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   6
      Text            =   "Size:"
      Top             =   2655
      Width           =   3345
   End
   Begin VB.TextBox txtMultipleSelectionText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   4
      Text            =   "items selected"
      Top             =   1860
      Width           =   3345
   End
   Begin VB.TextBox txtPrompt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   3
      Text            =   "Select an item to view its description."
      Top             =   1455
      Width           =   3345
   End
   Begin VB.TextBox txtFontName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   2
      Text            =   "verdana"
      Top             =   1065
      Width           =   3345
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   26
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Multiple File Delimiter:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   25
      Top             =   2295
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8520
      Y1              =   7530
      Y2              =   7530
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   21
      Top             =   3495
      Width           =   1005
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total File Size Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   20
      Top             =   3090
      Width           =   1905
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Size Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   19
      Top             =   2700
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Multiple Selection Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   18
      Top             =   1905
      Width           =   1785
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prompt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   17
      Top             =   1500
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3300
      TabIndex        =   16
      Top             =   1110
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":2261A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   210
      TabIndex        =   0
      Top             =   60
      Width           =   7905
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuInsertCalendar 
         Caption         =   "&Calendar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_FONT As String = "verdana"
Private Const DEFAULT_PROMPT_TEXT As String = "Select an item to view its description."
Private Const DEFAULT_MULTIPLE_TEXT  As String = " items selected."
Private Const DEFAULT_SIZE_TEXT As String = "Size: "
Private Const DEFAULT_FILE_SIZE_TEXT As String = "Total File Size: "
Private Const DEFAULT_DELIMITER_TEXT As String = ","
Private Const DEFAULT_BYTES_TEXT As String = "bytes"
Private Const DEFAULT_FOLDER_NAME As String = "%THISDIRNAME%"

'' used for finding the correct windows folder
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub ApplyChanges()
    '' the filenumber for our input data
    Dim iFileNumber As Integer
        
    '' the buffer to get our windows folder
    Dim strBuffer As String * 150, strData As String
    
    '' the path to our windows folder
    Dim strPath As String
    
    '' return value from api call
    Dim lReturnValue As Long
    
    '' call our api function
    lReturnValue = GetWindowsDirectory(strBuffer, Len(strBuffer))
    
    '' the path is the left$ the number of characters returned in lReturnValue
    If lReturnValue > 0 Then
        strPath = Left$(strBuffer, lReturnValue)
    Else
        MsgBox "Unable to located windows system folder.", vbCritical, "Unable to update"
        Exit Sub
    End If
    
    '' this is the file we are going to modify
    strPath = strPath & "\Web\Folder.htt"
    
    '' now load our data into a buffer
    '' make sure our data file exists
    If Dir$(App.Path & "\working.bak", vbHidden) <> vbNullString Then
        '' get a free file number
        iFileNumber = FreeFile
        
        '' open the file
        Open App.Path & "\working.bak" For Binary Access Read As iFileNumber
        
        '' reserve a buffer the size of our file
        strData = String(LOF(iFileNumber), 0)
        
        '' and read our data into the buffer
        Get #iFileNumber, , strData
        
        '' close our file
        Close iFileNumber
        
        '' now we add our custom information to the file
        strData = Replace$(strData, "#DEFAULT_PROMPT#", txtPrompt)
        strData = Replace$(strData, "#DEFAULT_FONT#", txtFontName)
        strData = Replace$(strData, "#DEFAULT_MULTIPLE_TEXT#", txtMultipleSelectionText)
        strData = Replace$(strData, "#DEFAULT_SIZE_TEXT#", txtSizeText)
        strData = Replace$(strData, "#DEFAULT_FILE_SIZE_TEXT#", txtTotalFileSizeText)
        strData = Replace$(strData, "#DEFAULT_DELIMITER_TEXT#", txtDelimiter)
        strData = Replace$(strData, "#DEFAULT_BYTES_TEXT#", txtBytesText)
        strData = Replace$(strData, "#MY_HTML_TEXT#", txtCustomHtml)
        strData = Replace$(strData, "#DEFAULT_FOLDER_NAME#", txtFolderName)
            
        '' now we have to replace the file with our data
        iFileNumber = FreeFile
        
        '' open the new one
        Open strPath For Output As iFileNumber
        
        '' write our data
        Print #iFileNumber, strData
        
        '' close the file
        Close iFileNumber
        
        '' Completed!
        MsgBox "Changes successful, click 'Refresh' or press F5 to refresh any open windows.", vbExclamation, "Success"
        
    End If
    
    
End Sub

Private Sub RestoreDefaults()
    txtBytesText = DEFAULT_BYTES_TEXT
    txtDelimiter = DEFAULT_DELIMITER_TEXT
    txtFolderName = DEFAULT_FOLDER_NAME
    txtFontName = DEFAULT_FONT
    txtMultipleSelectionText = DEFAULT_MULTIPLE_TEXT
    txtPrompt = DEFAULT_PROMPT_TEXT
    txtSizeText = DEFAULT_SIZE_TEXT
    txtTotalFileSizeText = DEFAULT_FILE_SIZE_TEXT
    txtCustomHtml = vbNullString

End Sub

Private Sub RestoreFolder()
    '' the buffer to get our windows folder
    Dim strBuffer As String * 150, strData As String
    
    '' the path to our windows folder
    Dim strPath As String
    
    '' long return value of api call
    Dim lReturnValue As Long
    
    '' make the api call
    lReturnValue = GetWindowsDirectory(strBuffer, Len(strBuffer))
    
    '' get our path
    strPath = Left$(strBuffer, lReturnValue)
    
    strPath = strPath & "\Web"
    
    '' copy our original file there
    If Dir$(strPath & "\Folder.htt", vbHidden) <> vbNullString And _
        Dir$(App.Path & "\original.bak") <> vbNullString Then
    
        FileCopy App.Path & "\original.bak", strPath & "\Folder.htt"
        MsgBox "Restoration Successful!", vbExclamation, "Success"
    Else
        MsgBox "Unable to restore", vbCritical, "Failure"
    End If

    
End Sub

Private Sub cmdApply_Click()
    '' prompt to make sure they want to change the system
    If MsgBox("Are you sure you want to make these system changes?", _
        vbQuestion Or vbYesNoCancel, "Change System?") = vbYes Then
        
        Call ApplyChanges
    End If
        
End Sub



Private Sub cmdCancel_Click()
    End
    
End Sub

Private Sub cmdDefaults_Click()
    Call RestoreDefaults
    
End Sub


Private Sub cmdInsertObject_Click()
    PopupMenu mnuPopup
    
End Sub

Private Sub cmdOK_Click()
    '' prompt to make sure they want to change the system
    If MsgBox("Are you sure you want to make these system changes?", _
        vbQuestion Or vbYesNoCancel, "Change System?") = vbYes Then
        
        Call ApplyChanges
        Unload Me
        
    End If

End Sub

Private Sub cmdRestore_Click()
    Call RestoreFolder
    
End Sub


Private Sub Form_Load()
    Call RestoreDefaults
    
End Sub



Private Sub mnuInsertCalendar_Click()
    Dim strTemp As String
    
    strTemp = "<object classid='clsid:232E456A-87C3-11D1-8BE3-0000F8754DA1' id='MonthView1' width='285' height='258'>" & _
                "<param name='_ExtentX' value='7541'>" & _
                "<param name='_ExtentY' value='6826'>" & _
                "<param name='_Version' value='393216'>" & _
                "<param name='ForeColor' value='0'>" & _
                "<param name='BackColor' value='16777215'>" & _
                "<param name='BorderStyle' value='0'>" & _
                "<param name='Appearance' value='1'>" & _
                "<param name='MousePointer' value='0'>" & _
                "<param name='Enabled' value='1'>" & _
                "<param name='OLEDropMode' value='0'>" & _
                "<param name='MaxSelCount' value='7'>" & _
                "<param name='MonthColumns' value='1'>" & _
                "<param name='MonthRows' value='1'>" & _
                "<param name='MonthBackColor' value='-2147483643'>" & _
                "<param name='MultiSelect' value='0'>" & _
                "<param name='ScrollRate' value='0'>" & _
                "<param name='ShowToday' value='1'>" & _
                "<param name='ShowWeekNumbers' value='0'>" & _
                "<param name='StartOfWeek' value='662831105'>" & _
                "<param name='TitleBackColor' value='-2147483633'>" & _
                "<param name='TitleForeColor' value='-2147483630'>" & _
                "<param name='TrailingForeColor' value='-2147483631'>" & _
                "<param name='CurrentDate' value='37250'>" & _
                "<param name='MaxDate' value='2958465'><param name='MinDate' value='-53688'></object>"
    txtCustomHtml = txtCustomHtml & vbCrLf & strTemp
    

End Sub


