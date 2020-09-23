VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInventory 
   BorderStyle     =   0  'None
   Caption         =   "Inventory Tracker"
   ClientHeight    =   8010
   ClientLeft      =   5280
   ClientTop       =   5655
   ClientWidth     =   13050
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInventory.frx":0000
      Left            =   1440
      List            =   "frmInventory.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   720
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search Results"
      Height          =   4215
      Left            =   10440
      TabIndex        =   38
      Top             =   1800
      Width           =   2535
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   21
         ToolTipText     =   "Click To Cancel Search: Edit/Delete"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.ListBox lstResults 
         Height          =   2460
         ItemData        =   "frmInventory.frx":0094
         Left            =   120
         List            =   "frmInventory.frx":0096
         TabIndex        =   18
         ToolTipText     =   "Click On The Record To Display"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblSearchMode 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   2880
         Width           =   2055
      End
   End
   Begin VB.ComboBox cboFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInventory.frx":0098
      Left            =   1440
      List            =   "frmInventory.frx":00A2
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Choose Which Field To Search From"
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      ToolTipText     =   "Click To Find Record"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   16
      ToolTipText     =   "Enter Search Criteria"
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Top             =   6840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdbOpen 
      Left            =   9360
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Spare Equipment (Max. 250 Characters)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4680
      TabIndex        =   35
      Top             =   3840
      Width           =   5655
      Begin VB.TextBox txtSEquip 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         ToolTipText     =   "Enter Spare Equipment (250 Characters Max)"
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Equipment (Max. 250 Characters)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4680
      TabIndex        =   25
      Top             =   1800
      Width           =   5655
      Begin VB.TextBox txtEquip 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         ToolTipText     =   "Enter Equipment (250 Characters Max)"
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Move To Last Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   24
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "Move To First Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   23
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   13
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Save Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      ToolTipText     =   "CLICK HERE TO SAVE NEW RECORD OR SAVE CHANGES"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add  New Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtPDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   9
      ToolTipText     =   "Enter Purchase Date"
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox txtPONum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   8
      ToolTipText     =   "Enter PO Number"
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox txtVendor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   7
      ToolTipText     =   "Enter Vendor"
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox txtSerialN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   6
      ToolTipText     =   "Enter Computer Serial Number"
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox txtNetPort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   5
      ToolTipText     =   "Enter Network Port"
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtCompN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   4
      ToolTipText     =   "Enter Computer Name"
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtDept 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   3
      ToolTipText     =   "Enter Department"
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtFirstN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   2
      ToolTipText     =   "Enter First Name"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtLastN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   1
      ToolTipText     =   "Enter Last Name (Required)"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblLastFirst 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   240
      TabIndex        =   43
      Top             =   1200
      Width           =   5955
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      Caption         =   "Sort By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9480
      TabIndex        =   40
      Top             =   120
      Width           =   3210
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Search For:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Purchase Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "PO Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Vendor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Serial Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Network Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Computer Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Department:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbInventory As Database
Dim rsTracker As Recordset
Dim SQLQuery As String
Dim Choice As String
Dim SQLSearch As String
Dim EditEnabled As Integer
Dim DeleteEnabled As Integer
Dim dbLastName As String

Private Sub EnableFields()
txtLastN.Enabled = True
txtFirstN.Enabled = True
txtDept.Enabled = True
txtCompN.Enabled = True
txtNetPort.Enabled = True
txtSerialN.Enabled = True
txtVendor.Enabled = True
txtPONum.Enabled = True
txtPDate.Enabled = True
txtEquip.Enabled = True
txtSEquip.Enabled = True
End Sub

Private Sub DisableFields()
txtLastN.Enabled = False
txtFirstN.Enabled = False
txtDept.Enabled = False
txtCompN.Enabled = False
txtNetPort.Enabled = False
txtSerialN.Enabled = False
txtVendor.Enabled = False
txtPONum.Enabled = False
txtPDate.Enabled = False
End Sub

Private Sub ClearFields()
txtLastN = ""
txtFirstN = ""
txtDept = ""
txtCompN = ""
txtNetPort = ""
txtSerialN = ""
txtVendor = ""
txtPONum = ""
txtPDate = ""
txtEquip = ""
txtSEquip = ""
End Sub

Private Sub DisableNavigation()
cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
End Sub

Private Sub EnableNavigation()
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
End Sub

Private Sub DisableEdit()
cmdAdd.Enabled = False
cmdUpdate.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False
End Sub

Private Sub EnableEdit()
cmdAdd.Enabled = True
cmdUpdate.Enabled = True
cmdEdit.Enabled = True
cmdDelete.Enabled = True
End Sub

Private Sub DisplayFields()
'On Error Resume Next
txtLastN = rsTracker!LastName
txtFirstN = rsTracker!FirstName
txtDept = rsTracker!Department
txtCompN = rsTracker!ComputerName
txtNetPort = rsTracker!NetworkPort
txtSerialN = rsTracker!SerialNumber
txtVendor = rsTracker!Vendor
txtPONum = rsTracker!PONumber
txtPDate = rsTracker!PurchaseDate
txtEquip = rsTracker!Equipment
txtSEquip = rsTracker!SpareEquipment
End Sub

Private Sub dbRefresh()
cmdNext.Enabled = True
cmdLast.Enabled = True
Call DisableFields
Call ClearFields
Set rsTracker = dbInventory.OpenRecordset(SQLQuery)
Call DisplayFields
End Sub

Private Sub dbUpdate()
With rsTracker
    !LastName = txtLastN
    !FirstName = txtFirstN
    !Department = txtDept
    !ComputerName = txtCompN
    !NetworkPort = txtNetPort
    !SerialNumber = txtSerialN
    !Vendor = txtVendor
    !PONumber = txtPONum
    !PurchaseDate = txtPDate
    !Equipment = txtEquip
    !SpareEquipment = txtSEquip
    .Update
End With
End Sub

Private Sub dbPopulateList()
lstResults.Clear
Do Until rsTracker.NoMatch = True
    If rsTracker.Fields(Choice) = "" Then
        rsTracker.FindNext SQLSearch
    Else
        lstResults.AddItem rsTracker.Fields(Choice)
        rsTracker.FindNext SQLSearch
    End If
Loop
End Sub

Private Sub ErrHandle()
ErrorHandler:
If Err = 3021 Then
    Exit Sub
End If
End Sub

Private Sub cboSort_Click()
On Error Resume Next
Dim Sort As String
Select Case cboSort.ListIndex
    Case 0
        Sort = "LastName"
    Case 1
        Sort = "FirstName"
    Case 2
        Sort = "Department"
    Case 3
        Sort = "ComputerName"
    Case 4
        Sort = "NetworkPort"
    Case 5
        Sort = "SerialNumber"
    Case 6
        Sort = "Vendor"
    Case 7
        Sort = "PONumber"
    Case 8
        Sort = "PurchaseDate"
End Select
SQLQuery = "SELECT * FROM Tracker WHERE " & Sort & " LIKE '*' ORDER BY " & Sort & ""
dbRefresh
End Sub

Private Sub cmdAdd_Click()
Call EnableFields
Call ClearFields
Call DisableEdit
lblLastFirst.Caption = ""
cmdSearch.Enabled = True
txtSearch.Enabled = True
cmdUpdate.Enabled = True
txtLastN.SetFocus
rsTracker.AddNew
End Sub

Private Sub cmdCancel_Click()
Call EnableEdit
Call dbRefresh
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
cmdPrevious.Enabled = False
cmdFirst.Enabled = False
cmdSearch.Enabled = True
lblSearchMode.Caption = ""
End Sub

Private Sub cmdDelete_Click()
Dim YesNo
On Error GoTo ErrorH
YesNo = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Delete")
If YesNo = vbNo Then Exit Sub
rsTracker.Delete
lstResults.Clear
If DeleteEnabled = 1 Then
    DeleteEnabled = 0
    Call dbRefresh
    Call dbPopulateList
    cmdSearch.Enabled = True
    cmdCancel.Enabled = False
    cmdAdd.Enabled = True
    MsgBox "Record deleted.", , "Delete"
    lblSearchMode.Caption = ""
    rsTracker.MoveFirst
    Call DisplayFields
    lblLastFirst.Caption = txtLastN & ", " & txtFirstN
    Exit Sub
End If
MsgBox "Moving to next Record.", , "Record Deleted"
rsTracker.MoveNext
If rsTracker.EOF Then
    MsgBox "You have deleted the last record...Moving to last available record.", , "Record Deleted"
    rsTracker.MovePrevious
    Call ClearFields
End If
Call DisplayFields
lblLastFirst.Caption = txtLastN & ", " & txtFirstN
Exit Sub
ErrorH:
MsgBox "All records have been deleted, database is empty!", , "No Records"
Call DisableEdit
Call DisableNavigation
Call ClearFields
Call ErrHandle
cmdAdd.Enabled = True
cmdSearch.Enabled = False
cmdCancel.Enabled = False
txtSearch.Enabled = False
lstResults.Clear
lblLastFirst.Caption = ""
End Sub

Private Sub cmdEdit_Click()
If txtLastN.Text = "" Then
    MsgBox "There are no records to edit.", , "No Records"
    Exit Sub
ElseIf EditEnabled = 2 Then
    Call EnableFields
    Call DisableEdit
    cmdUpdate.Enabled = True
    txtLastN.SetFocus
Else
    Call EnableFields
    Call DisableEdit
    dbLastName = txtLastN
    cmdUpdate.Enabled = True
    txtLastN.SetFocus
    EditEnabled = 1
End If
On Error GoTo ErrorH
rsTracker.Edit
ErrorH:
Call ErrHandle
End Sub

Private Sub cmdExit_Click()
Dim YesNo
YesNo = MsgBox("Are you sure you want to quit?", vbYesNo, "Exit")
If YesNo = vbNo Then Exit Sub
MsgBox "Make sure you backup the database file!!!", , "Backup Warning!!!!"
Unload Me
End Sub

Private Sub cmdFirst_Click()
On Error GoTo ErrorH
rsTracker.MoveFirst
Call EnableEdit
cmdUpdate.Enabled = False
Call DisableFields
Call DisableNavigation
cmdLast.Enabled = True
cmdNext.Enabled = True
If rsTracker.BOF Then
    rsTracker.MoveNext
End If
Call DisplayFields
lblLastFirst.Caption = txtLastN & ", " & txtFirstN
Exit Sub
ErrorH:
MsgBox "There are no records to display.", , "No Records"
Call ErrHandle
End Sub

Private Sub cmdLast_Click()
On Error GoTo ErrorH
rsTracker.MoveLast
Call EnableEdit
cmdUpdate.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdPrevious.Enabled = True
cmdFirst.Enabled = True
Call DisableFields
If rsTracker.EOF Then
    rsTracker.MovePrevious
End If
Call DisplayFields
lblLastFirst.Caption = txtLastN & ", " & txtFirstN
Exit Sub
ErrorH:
MsgBox "There are no records to display.", , "No Records"
Call ErrHandle
End Sub

Private Sub cmdNext_Click()
On Error GoTo ErrorH
rsTracker.MoveNext
Call EnableEdit
cmdUpdate.Enabled = False
cmdPrevious.Enabled = True
cmdFirst.Enabled = True
Call DisableFields
If rsTracker.EOF Then
    rsTracker.MovePrevious
    MsgBox "You are at the last record.", , "Last Record"
    cmdNext.Enabled = False
    cmdLast.Enabled = False
End If
Call DisplayFields
lblLastFirst.Caption = txtLastN & ", " & txtFirstN
Exit Sub
ErrorH:
MsgBox "There are no records to display.", , "No Records"
Call ErrHandle
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo ErrorH
rsTracker.MovePrevious
Call EnableEdit
cmdUpdate.Enabled = False
cmdNext.Enabled = True
cmdLast.Enabled = True
Call DisableFields
If rsTracker.BOF Then
    rsTracker.MoveNext
    MsgBox "You are at the first record.", , "First Record"
    cmdPrevious.Enabled = False
    cmdFirst.Enabled = False
End If
Call DisplayFields
lblLastFirst.Caption = txtLastN & ", " & txtFirstN
Exit Sub
ErrorH:
MsgBox "There are no records to display.", , "No Records"
Call ErrHandle
End Sub

Private Sub cmdSearch_Click()
lstResults.Clear
If cboFields = "Serial Number" Then
    Choice = "SerialNumber"
Else
    Choice = "PONumber"
End If
SQLSearch = Choice & " LIKE '" & txtSearch & "'"
rsTracker.FindFirst SQLSearch
If rsTracker.NoMatch Then
    MsgBox "Try using a wild card such as " & txtSearch & "* or *" & txtSearch, , "No Record Found"
    Exit Sub
End If
Call dbPopulateList
Call DisableNavigation
Call dbRefresh
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrorH
If txtLastN = "" Then
    MsgBox "Please enter last name.", , "Required Field"
    txtLastN.SetFocus
    Exit Sub
End If
Call dbUpdate
If EditEnabled = 1 Then
    EditEnabled = 0
    Call EnableEdit
    Call DisableFields
    If dbLastName <> txtLastN Then Call dbRefresh
    cmdUpdate.Enabled = False
    lblLastFirst.Caption = txtLastN & ", " & txtFirstN
ElseIf EditEnabled = 2 Then
    EditEnabled = 0
    Call EnableEdit
    Call dbRefresh
    Call dbPopulateList
    rsTracker.MoveFirst
    Call DisplayFields
    lblLastFirst.Caption = txtLastN & ", " & txtFirstN
    cmdSearch.Enabled = True
    cmdUpdate.Enabled = False
    cmdCancel.Enabled = False
    lblSearchMode.Caption = ""
Else
    MsgBox "Your record has been saved.", , "Update"
    Call EnableEdit
    Call dbRefresh
    rsTracker.MoveLast
    Call DisplayFields
    lblLastFirst.Caption = txtLastN & ", " & txtFirstN
    cmdUpdate.Enabled = False
    cmdPrevious.Enabled = True
    cmdFirst.Enabled = True
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    Exit Sub
End If
MsgBox "Your record has been saved.", , "Update"
Call EnableNavigation
Exit Sub
ErrorH:
Call ErrHandle
End Sub

Private Sub Form_Load()
Call DisableEdit
Call DisableNavigation
Call DisableFields
cboFields.ListIndex = 0
Label12.Caption = "Database Not Open!"
txtEquip.Enabled = False
txtSEquip.Enabled = False
txtSearch.Enabled = False
cmdSearch.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub lstResults_Click()
Dim SQLResults As String
SQLResults = "SELECT * FROM Tracker WHERE " & Choice & " LIKE '" & lstResults & "' ORDER BY LastName, FirstName"
Set rsTracker = dbInventory.OpenRecordset(SQLResults)
Call DisplayFields
Call DisableNavigation
cmdAdd.Enabled = False
cmdCancel.Enabled = True
cmdSearch.Enabled = False
EditEnabled = 2
DeleteEnabled = 1
lblSearchMode.Caption = "Search: Edit/Delete Mode Activated"
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
Dim YesNo
YesNo = MsgBox("Are you sure you want to quit?", vbYesNo, "Exit")
If YesNo = vbNo Then Exit Sub
MsgBox "Make sure you backup the database file!!!", , "Backup Warning!!!!"
Unload Me
End Sub

Private Sub mnuOpen_Click()
Dim dataPath As String
On Error GoTo ErrorH
cdbOpen.Filter = "Access 2002 Files (*.MDB)|*.MDB"
cdbOpen.ShowOpen
dataPath = cdbOpen.FileName
If dataPath = "" Then
    Exit Sub
End If
Set dbInventory = OpenDatabase(dataPath)
SQLQuery = "SELECT * FROM Tracker WHERE LastName LIKE '*' ORDER BY LastName, FirstName"
Set rsTracker = dbInventory.OpenRecordset(SQLQuery)
Call DisplayFields
Call EnableEdit
lblLastFirst.Caption = txtLastN & ", " & txtFirstN
Label12.Caption = "Database Is Open!"
cboSort.ListIndex = 0
cboSort.Enabled = True
cmdUpdate.Enabled = False
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdSearch.Enabled = True
txtSearch.Enabled = True
Exit Sub
ErrorH:
If Err = 3343 Or Err = 3078 Then
    MsgBox "The database you are trying to open is not a valid database and/or file!", , "Database Error"
    Exit Sub
ElseIf txtLastN.Text = "" Then
    cboSort.Enabled = True
    cboSort.ListIndex = 0
    MsgBox "The database is empty.", , "No Records"
    Call DisableEdit
    Call DisableNavigation
    Label12.Caption = "Database Is Open!"
    cmdAdd.Enabled = True
    cmdSearch.Enabled = False
    txtSearch.Enabled = False
End If
End Sub

