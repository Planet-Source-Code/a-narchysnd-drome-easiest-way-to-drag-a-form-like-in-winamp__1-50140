VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simplest and Fastest way to Drag a Form!
'Written by Darwin Yu - Copyright 2003
'You can even drag the form by clicking on other objects on the screen...
'...just copy and paste the corresponding code into the object's code:
'For example, if you're trying to drag a form by clicking on a Picturebox,
'Copy the Form_Mousedown code into Picture1_Mousedown
'And do so with the other 2...its that easy!
'Please vote if you like it!
'No API calls, no nothing.

''''
Dim sL As Integer 'Starting Left coord
Dim eL As Integer 'Ending Left coord
Dim totalL As Integer 'difference of endingleft - startingleft = amount actually moved
Dim sT As Integer 'starting top coord
Dim eT As Integer 'ending top coord
Dim totalT As Integer 'difference of endingtop - startingtop = amount actually moved
Dim LT As Integer 'Holds the steps
''''

Private Sub Form_Load()
MsgBox "click anywhere on form to drag!"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LT = 0 Then 'if user first clicked, then
sL = X
sT = Y
LT = 1
End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If LT = 1 Then 'if user hasnt let the mousebutton go, then
eL = X
eT = Y '
totalL = eL - sL
totalT = eT - sT
Form1.Left = Form1.Left + totalL
Form1.Top = Form1.Top + totalT

End If
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LT = 0 'user has let mouse button go, therefore, reset
End Sub
