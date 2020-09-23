VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmIndexServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index Server Search"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   Icon            =   "frmIndexServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCVtext 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      TabIndex        =   19
      Top             =   180
      Width           =   2655
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help!"
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame frOptions 
      Caption         =   "Index Server Criteria"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   5055
      Begin VB.TextBox txtHostName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1020
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         ToolTipText     =   "Host Name or Machine Name that the Index Server Service is running on..."
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optANDOR 
         Caption         =   "OR"
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   5
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox txtReturn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "50"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtReturned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox cboSort 
         Height          =   315
         ItemData        =   "frmIndexServer.frx":57E2
         Left            =   1020
         List            =   "frmIndexServer.frx":57F5
         TabIndex        =   8
         Text            =   "Rank"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtCatalog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1020
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "System"
         ToolTipText     =   "Catalog name..."
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optANDOR 
         Caption         =   "AND"
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   6
         Top             =   1140
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Host Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Results to Return:"
         Height          =   315
         Left            =   2940
         TabIndex        =   14
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Results Returned:"
         Height          =   315
         Left            =   2940
         TabIndex        =   13
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sort By:"
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Catalog:"
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Top             =   780
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdSearchCV 
      Caption         =   "Search"
      Height          =   375
      Left            =   6900
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCVtext 
      Caption         =   ">"
      Height          =   435
      Index           =   0
      Left            =   5280
      TabIndex        =   2
      Top             =   180
      Width           =   435
   End
   Begin VB.CommandButton cmdCVtext 
      Caption         =   "<"
      Height          =   435
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      Top             =   720
      Width           =   435
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      Top             =   2940
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwWord 
      Height          =   2295
      Left            =   5760
      TabIndex        =   15
      Top             =   180
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   4048
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Search Word"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmIndexServer.frx":581C
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmIndexServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCVtext_Click(Index As Integer)
If Index = 0 Then
    If txtCVtext.Text <> "" Then
        lvwWord.ListItems.Add , , txtCVtext.Text
    End If
Else
    If lvwWord.ListItems.Count > 0 Then
        lvwWord.ListItems.Remove lvwWord.SelectedItem.Index
    End If
End If
End Sub

Private Sub cmdHelp_Click()
frmHelp.Show
End Sub

Private Sub cmdSearchCV_Click()
    IndexServerSearch
End Sub

Private Sub lvwResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If lvwResults.ListItems.Count > 0 Then
    
    Static lngLastColumnSorted As Long
    
    'When a ColumnHeader object is clicked, the ListView control is
    'sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader -1
    lvwResults.SortKey = ColumnHeader.Index - 1
    
    'if the last column sorted was the same, then switch sort order
    ' otherwise sort new chosen column ascending
    If lngLastColumnSorted = ColumnHeader.Index - 1 Then
        If lvwResults.SortOrder = lvwAscending Then
            lvwResults.SortOrder = lvwDescending ' sort descending
        Else
            lvwResults.SortOrder = lvwAscending ' sort order ascending
        End If
    Else
        lvwResults.SortOrder = lvwAscending ' sort order ascending
    End If
    
    'remember last column sorted
    lngLastColumnSorted = ColumnHeader.Index - 1
    
    'set sorted to True to sort the list.
    lvwResults.Sorted = True
    
    'ensure no records are selected
    lvwResults.SelectedItem.Selected = False
End If
End Sub

Private Sub lvwResults_DblClick()

With lvwResults.ListItems
    MsgBox "The item selected is located at;" & vbNewLine & vbNewLine & _
        .Item(lvwResults.SelectedItem.Index).ListSubItems(5).Text, vbInformation, "File Location"
End With

End Sub

Private Sub txtCVtext_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then 'Return key
    cmdCVtext_Click 0
End If

End Sub

Public Function IndexServerSearch()

Dim oRS                     As ADODB.Recordset
Dim oQuery                  As CissoQuery
Dim objUtil                 As CissoUtil
Dim sSearchString           As String
Dim sAndOr                  As String
Dim i                       As Integer

On Error GoTo Err_Handler

'If no Host Name was specified...
If txtHostName.Text = "" Then
   MsgBox "Please enter the computer name (Host Name) that you want to search.", vbExclamation, "Host Name"
   txtHostName.SetFocus
   Exit Function
End If

'If no Catalog was specified
If txtCatalog.Text = "" Then
   MsgBox "Please enter the Catalog that you want to search.", vbExclamation, "Catalog"
   txtHostName.SetFocus
   Exit Function
End If

For i = 1 To lvwWord.ListItems.Count
    If i = 1 Then
        sSearchString = Trim(lvwWord.ListItems(i).Text)
    Else
        sSearchString = sSearchString & sAndOr & Trim(lvwWord.ListItems(i).Text)
    End If
Next

'If no search term was specified, then squawk!
If sSearchString = "" Then
   MsgBox "Please enter words to search on.", vbExclamation, "Search Word(s)"
   txtCVtext.SetFocus
   Exit Function
End If

'And Or
For i = 0 To optANDOR.Count - 1
    If optANDOR(i).Value = True Then
        sAndOr = " " & optANDOR(i).Caption & " "
    End If
Next

'The path to the files to be searched remember to append a
'wildcard * to the end or beginning of this constant as appropriate..

'A search term was specified, so show the search results
If sSearchString <> "" Then
    'query://hostname/indexname
    Set oQuery = New CissoQuery   'Server.CreateObject("ixsso.Query")
    
    'Build the search query
    oQuery.Query = "CONTAINS " & sSearchString
    
    'The maximum number of records to be returned..
    oQuery.MaxRecords = txtReturn.Text
    
    'Sort the results..
    oQuery.SortBy = cboSort.Text '"filename[d]"
    
    'Specify which columns are returned..
    'oQuery.Columns = "vpath,path,filename,size,write,characterization"
    oQuery.Columns = "DocTitle,FileName,Rank,Write,Size,Path" '
    
    'Indicate which catalog and hostname to use...
    oQuery.Catalog = "query://" & Trim$(txtHostName.Text) & "/" & Trim$(txtCatalog.Text) 'WK007630

    
    'Set objUtil = CreateObject("ixsso.util") 'New CissoUtil
    'objUtil.AddScopeToQuery oQuery, "C:\Documents and Settings\zzGStevens\My Documents\Example Code\IndexServer", "shallow"

    
    'Create the results RecordSet
    Set oRS = oQuery.CreateRecordset("nonsequential")

    'Count the number of results returned
    If Err.Number = 0 Then
        txtReturned = oRS.RecordCount
        If Not oRS.EOF Then
            BuildLVW oRS, lvwResults
        Else
            MsgBox "No results were returned from your query.", vbInformation, "Query"
        End If
    End If
End If

Err_Handler:

    Set oQuery = Nothing
    If Err <> 0 Then
        MsgBox "The search gave the following error: " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation
    End If
    
End Function


Private Function BuildLVW(rsListView As ADODB.Recordset, oListView As ListView, _
    Optional bLeaveList As Boolean)
'<EhHeader>

   On Error GoTo BuildLVW_Err
'</EhHeader>

Dim chColHead       As ColumnHeader
Dim itmNewLine      As ListItem
Dim intIndex        As Integer
Dim intCount2       As Integer
Dim intTotCount     As Integer
Dim isOkay          As Boolean
Dim sColumnHeader   As String
Dim fWidth          As Single

    With rsListView
        'Clear out the ListView ready for new values....
        oListView.ColumnHeaders.Clear
        If bLeaveList = False Then oListView.ListItems.Clear
        'Set up the column headers and then add the items to the ListView control....
        
        'First, Add the Column Headers
        For intIndex = 0 To .Fields.Count - 1
            Set chColHead = oListView.ColumnHeaders.Add(, , rsListView(intIndex).Name)
            fWidth = Len(rsListView(intIndex).Name) * 130
            oListView.ColumnHeaders.Item(intIndex + 1).Width = fWidth 'Adjust size if needed...
        Next intIndex
        
        ' Now, loop through the recordset and add Items to the lvw.....
        If .EOF = False Then
            intTotCount = rsListView.RecordCount
            For intIndex = 1 To intTotCount
                If IsNull(rsListView(1).Value) = False Then
                    Set itmNewLine = oListView.ListItems.Add(, , Trim$(rsListView(1).Value))
                    
                    For intCount2 = 1 To .Fields.Count - 1 'Add the subitems..
                        If IsNull(rsListView(intCount2).Value) Then
                            itmNewLine.SubItems(intCount2) = ""
                        Else
                            itmNewLine.SubItems(intCount2) = Trim$(rsListView(intCount2).Value)
                        End If
                        If intIndex <= oListView.ColumnHeaders.Count Then
                            If oListView.ColumnHeaders.Item(intCount2 + 1).Width < Len(Trim$(rsListView(intCount2).Value)) * 110 Then
                                oListView.ColumnHeaders.Item(intCount2 + 1).Width = Len(Trim$(rsListView(intCount2).Value)) * 110 'Adjust size if needed...
                            End If
                        End If
                    Next intCount2
                    .MoveNext
                Else
                    Set itmNewLine = oListView.ListItems.Add(, , "")
                End If
            Next intIndex
            oListView.GridLines = True
        Else
            oListView.ListItems.Clear
            oListView.GridLines = False
        End If
    End With

'<EhFooter>
BuildLVW_Exit:

Exit Function
BuildLVW_Err:

    If Err <> 0 Then
        MsgBox "The search gave the following error in function BuildLVW: " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation
    End If

End Function

