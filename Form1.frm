VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Interactions"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB 
      Height          =   315
      Left            =   1350
      TabIndex        =   13
      Top             =   7620
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   90
      TabIndex        =   2
      Top             =   2160
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Images List"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtContent"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Forms"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lvFormElements"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstForms"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtFormInputs"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtSelected"
      Tab(2).Control(1)=   "lblStuff"
      Tab(2).Control(2)=   "Label1(4)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Tree"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tvList"
      Tab(3).Control(1)=   "txtHTML"
      Tab(3).ControlCount=   2
      Begin VB.TextBox txtHTML 
         Height          =   4965
         Left            =   -71610
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   360
         Width           =   6465
      End
      Begin VB.TextBox txtSelected 
         Height          =   2985
         Left            =   -74910
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   720
         Width           =   9675
      End
      Begin VB.TextBox txtFormInputs 
         Height          =   1335
         Left            =   2550
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3540
         Width           =   7155
      End
      Begin VB.TextBox txtContent 
         Height          =   4875
         Left            =   -74940
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   390
         Width           =   9705
      End
      Begin VB.ListBox lstForms 
         Height          =   2595
         Left            =   120
         TabIndex        =   3
         Top             =   690
         Width           =   2325
      End
      Begin MSComctlLib.ListView lvFormElements 
         Height          =   2625
         Left            =   2520
         TabIndex        =   6
         Top             =   690
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   4630
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Id"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Value"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.TreeView tvList 
         Height          =   4935
         Left            =   -74940
         TabIndex        =   11
         Top             =   360
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   8705
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label lblStuff 
         Caption         =   "Label2"
         Height          =   915
         Left            =   -74850
         TabIndex        =   16
         Top             =   3870
         Width           =   4605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Element Source HTML"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2550
         TabIndex        =   14
         Top             =   3330
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Any Text you have selected on the Web page will show here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   -74940
         TabIndex        =   9
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Imput Elements in Form - Select Element to view Source"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   4770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forms List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   450
         Width           =   870
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9570
      Top             =   1650
   End
   Begin VB.ListBox lstWindows 
      Height          =   1620
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   8520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Building Tree"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   7620
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of Open Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   180
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iSelected As Long
Private WithEvents Explorer As SHDocVw.ShellWindows ' Expose the Explorer Windows Events
Attribute Explorer.VB_VarHelpID = -1
Private WithEvents Document As HTMLDocument ' Expose the Selected documents events
Attribute Document.VB_VarHelpID = -1

Private Function Document_onclick() As Boolean
      Debug.Print "Left Mouse Clicked "
End Function
Private Function Document_oncontextmenu() As Boolean
      Debug.Print "Right Mouse Clicked "
End Function
Private Sub Document_onfocusin()
      Debug.Print "Document_onfocusin" '
End Sub
Private Sub Document_onfocusout()
      Debug.Print "Document_onfocusout" ' Window changed
End Sub
Private Sub Document_onmouseup()
Dim oSelect As IHTMLSelectionObject ' Holds a select field
      ' Set a reference to the select object, this will allow
      ' you to retrieve all text selected on the window
      ' we are linked too
      Set oSelect = Document.selection
      If oSelect.Type <> "None" Then
         ' The Selection object contains selected text
         Me.txtSelected.Text = oSelect.createRange.Text
      End If
      Debug.Print "Document_onmouseup"
End Sub

Private Sub Explorer_WindowRegistered(ByVal lCookie As Long)
      ' A new window was created and opened
      ' Set the timer to fire in 1/2 second to refresh the open windows list
      ' we have to wait for this event to finish before we can get
      ' some information from the new window
      
      Me.Timer1.Interval = 500
      Me.Timer1.Enabled = True
      Debug.Print Explorer.Count
End Sub

Private Sub Explorer_WindowRevoked(ByVal lCookie As Long)
      ' a window was closed so lets refresh the windows list
      ' and remove it
      Debug.Print Explorer.Count
      GetList
End Sub

Private Sub Form_Load()
      frmAbout.Show 1
      
      GetList ' Guild a list of open explorer windows

      Set Explorer = New SHDocVw.ShellWindows ' Setup a hook into the Explorer Windows Collection
      Debug.Print Explorer.Count ' Show how many Windows are currently created
      
      ' This timer is used to refresh the Window List every 5 Seconds
      Me.Timer1.Interval = 5000
      Me.Timer1.Enabled = True
End Sub

Sub GetList()
      ' Set the object (oSWin) ready to hold a Explorer Windows collection
      ' The ShellWindows object represents a collection of the open windows
      ' that belong to the shell. In fact, this collection contains references
      ' to Internet Explorer as well as other windows belonging to the shell,
      ' such as the Windows Explorer.
Dim oSWin As New SHDocVw.ShellWindows

      ' Set the object (oIE) ready to hold a Explorer Window
      ' This allows easy access to the methods and properties of that object
Dim oIE As SHDocVw.InternetExplorer


      lstWindows.Clear ' Clear down the Windows List
      
      ' We will step through each open window in the collection to find
      ' only windows that are on web sites, using the simplest method
      ' Check the URL starts with HTTP:// or HTTPS://
      ' Another method to use is as follows
      
      ' Dim oDoc As Variant
      ' For Each oIE In oSWin
      ' Set oDoc = oIE.document
      ' If TypeOf oDoc Is HTMLDocument Then
      ' lstWindows.AddItem oIE.LocationURL
      ' lstWindows.ItemData(lstWindows.NewIndex) = oIE.hWnd
      ' End If
      ' Next
      
      For Each oIE In oSWin

         If UCase$(oIE.LocationURL) Like "HTTP*://*" Then
            ' The LocationURL specifies the current Location the browser is pointing too
            lstWindows.AddItem oIE.LocationURL
            ' Record the window long handle for later use for when user clicks the
            ' window item so we can quickly get the that window object
            lstWindows.ItemData(lstWindows.NewIndex) = oIE.hWnd
         End If
      Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
      ' Tidy up
      Set Document = Nothing
      Set Explorer = Nothing
End Sub

Private Sub lstForms_Click()
Dim oDocument As HTMLDocument
Dim oForm As HTMLFormElement
Dim oFormInput As HTMLInputElement
Dim oFormSelect As HTMLSelectElement
Dim oFormArea As HTMLTextAreaElement
Dim oFormButton As HTMLInputButtonElement
Dim sNodeName As String
Dim iForm As Long
Dim i As Long
Dim ii As Long

      iForm = lstForms.ListIndex
      
      Set oDocument = GetActiveDocument(iSelected) ' Set a reference to the document

      Set oForm = oDocument.Forms.Item(iForm) ' Set a reference to the select form

      Me.lvFormElements.ListItems.Clear
      For i = 0 To oForm.length - 1

         sNodeName = oForm.Item(i).nodeName

         If UCase$(oForm.Item(i).Type) = "SUBMIT" Then
            sNodeName = "SUBMIT"
         End If

         Select Case sNodeName
            Case "INPUT"
               ' Display information about an input field
               ' These are hidden fields or text boxes, radio or options
               
               ' Set reference to the form input element
               ' Doing this allows up to easily see the properties and methods
               ' available
               Set oFormInput = oForm.Item(i)
               Me.lvFormElements.ListItems.Add , , oFormInput.Name
               With Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count)
                  .SubItems(1) = oFormInput.Type
                  .SubItems(2) = oFormInput.id
                  .SubItems(3) = oFormInput.Value
                  .SubItems(4) = oFormInput.outerHTML
                  If UCase$(oFormInput.Type) <> "HIDDEN" Then
                     '
                  Else
                     .ForeColor = &H808080
                     .ListSubItems.Item(1).ForeColor = &H808080
                     .ListSubItems.Item(2).ForeColor = &H808080
                     .ListSubItems.Item(3).ForeColor = &H808080
                     .ListSubItems.Item(4).ForeColor = &H808080
                        
                  End If
               End With
            Case "SELECT"
               ' Selection Input item
               Set oFormSelect = oForm.Item(i)

               Me.lvFormElements.ListItems.Add , , oFormSelect.Name
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(1) = "Select"
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(3) = oFormSelect.Value
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(2) = oFormSelect.id
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(4) = oFormSelect.outerHTML
            Case "TEXTAREA"
               ' Speaks for itself
               Set oFormArea = oForm.Item(i)
               Me.lvFormElements.ListItems.Add , , oFormArea.Name
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(1) = "TextArea"
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(3) = oFormArea.Value
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(2) = oFormArea.id
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(4) = oFormArea.outerHTML
            Case "SUBMIT"
               ' Button Item
               Set oFormButton = oForm.Item(i)
               Me.lvFormElements.ListItems.Add , , oFormButton.Name
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(1) = "Button"
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(3) = oFormButton.Value
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(2) = oFormButton.id
               Me.lvFormElements.ListItems.Item(Me.lvFormElements.ListItems.Count).SubItems(4) = oFormButton.outerHTML
               ' oFormButton.Value = "Test This Out"
               ' oFormButton.Click
         End Select
      Next
      
      ' Clean Up
      Set oDocument = Nothing
      Set oForm = Nothing
      Set oFormInput = Nothing
      Set oFormSelect = Nothing
      Set oFormArea = Nothing
      Set oFormButton = Nothing

End Sub
' ==============================================================
' Procedure:    GetActiveDocument
'
' Created on:   26 May 2005    By  Darren Lawrence
'
' Function :-
' Returns a Document Object using the hwnd to find it
'
' ==============================================================
Function GetActiveDocument(lHWND As Long) As HTMLDocument
      
      ' Set the object (oSWin) ready to hold a Explorer Windows collection
      ' The ShellWindows object represents a collection of the open windows
      ' that belong to the shell. In fact, this collection contains references
      ' to Internet Explorer as well as other windows belonging to the shell,
      ' such as the Windows Explorer.
Dim oSWin As New SHDocVw.ShellWindows

      ' Set the object (oIE) ready to hold a Explorer Window
      ' This allows easy access to the methods and properties of that object
Dim oIE As SHDocVw.InternetExplorer
Dim i As Long

      For i = 0 To oSWin.Count - 1 ' Loop through all open explorer windows
         If oSWin.Item(i).hWnd = lHWND Then
            ' This item in the windows list matches the one selected
            Set oIE = oSWin.Item(i) ' Set a reference to the selected window
            Set GetActiveDocument = oIE.Document ' Set a reference to the document
         End If
      Next i
      Set oIE = Nothing
      Set oSWin = Nothing
End Function
Private Sub lstWindows_Click()
Dim oDocument As HTMLDocument       ' Holds the Explorer Window content
Dim oImage As HTMLImg               ' Holds a explorer document image item
Dim oForm As HTMLFormElement        ' Holds a Form from a document
Dim oFormField As HTMLInputElement  ' Holds an Form Input Field
Dim oFormButton As HTMLInputButtonElement ' Holds a Form Button
Dim oSelect As IHTMLSelectionObject ' Holds a select form field
Dim oElement As IHTMLElement        ' Provides the ability to access the properties and methods that are common to all element objects
Dim i As Long
Dim ii As Long
Dim sName As String

      iSelected = lstWindows.ItemData(lstWindows.ListIndex) ' Get the selected windows handle
      
      Set oDocument = GetActiveDocument(iSelected) ' Set a reference to the document
            
      ' Set a reference to the select object, this will allow
      ' you to retrieve all text selected on the window
      ' we are linked too
      Set oSelect = oDocument.selection
      If oSelect.Type <> "None" Then
         ' The Selection object contains selected text
         Me.txtSelected.Text = oSelect.createRange.Text
      End If
                       
      Me.txtContent.Text = ""
      Me.lstForms.Clear
            
      ' Show images used
      ' Loop through the Images collection on the document
      For ii = 0 To oDocument.images.length - 1
         Set oImage = oDocument.images.Item(ii) ' Set a reference to the image
         ' Display the href of the image
         Me.txtContent.Text = Me.txtContent.Text & oImage.href & vbCrLf
      Next
      
      ' Display all Forms
      ' Loop through all Forms on this document
      
      For ii = 0 To oDocument.Forms.length - 1
         sName = oDocument.Forms.Item(ii).Name
         If sName = "" Then
            sName = "{no name}"
         End If
         lstForms.AddItem sName
      Next

      ' Build a tree list of the document elements

      Me.tvList.Nodes.Clear
      Me.tvList.Nodes.Add , , "r", "Document"
      Me.tvList.Nodes.Add "r", tvwChild, "Elements", "Elements"
      Me.tvList.Tag = iSelected ' Record for quick access if an element is selected by the user

      Me.PB.Max = oDocument.All.length - 1

      ' loop through all the documents elements
      
      For ii = 0 To oDocument.All.length - 1
         Me.PB.Value = ii
         Set oElement = oDocument.All.Item(ii) ' Set reference to the current element
         Me.tvList.Nodes.Add "Elements", tvwChild, "n" & Str(ii), oElement.nodeName
      Next
      Me.tvList.Nodes.Item(2).Expanded = True
      Me.tvList.Nodes.Item(2).EnsureVisible

      Set Document = oDocument ' hook into the current Documents Events
      ' Clean up
      Set oSelect = Nothing
      Set oDocument = Nothing
      Set oImage = Nothing
      Set oForm = Nothing
      Set oFormField = Nothing
      Set oFormButton = Nothing
      Set oElement = Nothing
End Sub

Private Sub lvFormElements_Click()
      On Error Resume Next
      Me.txtFormInputs.Text = Me.lvFormElements.SelectedItem.SubItems(4)
End Sub

Private Sub Timer1_Timer()
      Me.Timer1.Enabled = False
      GetList
      Me.Timer1.Interval = 5000
      Me.Timer1.Enabled = True
End Sub

Private Sub tvList_Click()
Dim i As Long
Dim oDocument As HTMLDocument
Dim oElement As IHTMLElement

      Me.txtHTML.Text = Me.tvList.SelectedItem.Tag
      Debug.Print Me.tvList.SelectedItem.Index - 3
      If Me.tvList.SelectedItem.Index - 3 > -1 Then
         Set oDocument = GetActiveDocument(Me.tvList.Tag) ' Set a reference to the document
         Set oElement = oDocument.All.Item(Me.tvList.SelectedItem.Index - 3)
         Me.txtHTML.Text = oElement.outerHTML
      End If

      ' Clean up
      Set oDocument = Nothing
      Set oElement = Nothing
End Sub
