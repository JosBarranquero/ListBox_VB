VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Students"
   ClientHeight    =   3105
   ClientLeft      =   8865
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemove 
      Appearance      =   0  'Flat
      TabStop         =   0  'False
      Height          =   375
      Left            =   4200
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox lstNames 
      Height          =   2400
      ItemData        =   "frmMain.frx":0442
      Left            =   120
      List            =   "frmMain.frx":0444
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Author:  José Antonio Barranquero Fernández
'Version: 1.0
'Date:    02/04/2017
'Remark:  Simple application that lets add and delete students in a list

Dim Student_File As String

'Subroutine (event) that occurs when the form opens for the first time
'We load the names from a file called "Students.txt" if it exists
Private Sub Form_Load()
    Let Student_File = App.Path & "\Students.txt"
    If Dir(Student_File) <> "" Then
        Dim curLine As String
        Dim iFileNo As Integer
        iFileNo = FreeFile
        Open Student_File For Input As #iFileNo
            Do While Not EOF(iFileNo)
                Line Input #iFileNo, curLine
                lstNames.AddItem curLine
            Loop
        Close #iFileNo
    Else
        Exit Sub
    End If
End Sub

'Subroutine (event) that occurs when the form is closing
'We save the names to a file called "Students.txt"
Private Sub Form_Unload(Cancel As Integer)
    If lstNames.ListCount <> 0 Then
        Dim iFileNo As Integer
        iFileNo = FreeFile
        Open Student_File For Output As #iFileNo
        For n = 0 To (lstNames.ListCount - 1)
            Print #iFileNo, lstNames.List(n)
        Next n
        Close #iFileNo
    Else
        If Dir(Student_File) <> "" Then Kill Student_File
    End If
End Sub

'Subroutine (event) that controls the Add button
Private Sub btnAdd_Click()
        Dim name As String
        name = txtName.Text       'We get the name from the textbox
        If checkName(name) Then         'Checking if the name is empty
            lstNames.AddItem name     'We add the name to our list
            txtName.Text = ""           'We clear the textbox input
        Else
            MsgBox "Input a name", vbCritical, "Empty name"
        End If
End Sub

'Subroutines (events) that controls the Remove button
Private Sub btnRemove_Click()
    Dim selected As Integer
    Let selected = lstNames.ListIndex
    If checkSelected(selected) Then             'We check if a name has been selected
        lstNames.RemoveItem selected          'We delete the item selected
    Else
        MsgBox "Select a name to remove", vbInformation, "Not selected"
    End If
End Sub

'Subroutine (event) that controls when a Key is pressed an Item is selected
Private Sub lstNames_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDelete) Then Call btnRemove_Click
End Sub

'Subroutine (event) that controls when a Key is pressed when the text field is focused
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Then Call btnAdd_Click
End Sub

'Function which checks if the name input is empty
Private Function checkName(name As String) As Boolean
        Dim valid As Boolean
        Let valid = False
        If Not (name = "") Then
            valid = True
        End If
        checkName = valid
End Function

'Function which checks if an item has been selected, -1 not selected, else is selected
Private Function checkSelected(selected As Integer) As Boolean
        Dim valid As Boolean
        Let valid = False
        If Not (selected = -1) Then
            valid = True
        End If
        checkSelected = valid
End Function
