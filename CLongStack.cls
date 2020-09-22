VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLongStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================================================================
' CLongStack
'---------------------------------------------------------------------------
'
' Name:           CLongStack
' Author:         Daniel Keep
' Contact:        shortcircuitfky@hotmail.com
'                 When contacting me, please include [VB] in the subject
'                 line, or I might not read your email.
' Version:        1.0
' Last Modified:  4 July 2002
' Description:    Implements a light-weight long stack class.
'
'===========================================================================

' NOTE: Some stack terminology for those who HAVEN'T done ASM ;)
'  Stack: A stack is a block that grows downwards in memory.  Essentially,
'         it's like a variable that has a 'memory' of it's own.  It
'         remembers all the previous values it's held.  To add to the top
'         of a stack, you 'push' new data on, and to remove the topmost
'         item, you 'pop' it off.  Just think of a stack of plates: last
'         one on is the first one off.
'  Frame: A frame is one of the plates on that stack.  Essentially, you can
'         only access the topmost frame of a stack at any one time.

' This is the stack itself
Private m_stack() As Long

' This is the 'height' of the stack.  0 indicates no frames
Private m_height&

'===========================================================================
' Init the height
Private Sub Class_Initialize()

m_height = 0&

End Sub

'===========================================================================
' Erase the stack
Private Sub Class_Terminate()

Erase m_stack
m_height = 0&

End Sub

'===========================================================================
' CLongStack.Push()
'   Pushes a new frame onto the stack
Public Sub Push(ByRef Frame&)

' Resize the stack array to accomodate the new frame
ReDim Preserve m_stack(m_height)

' Add the new frame
m_stack(m_height) = Frame
m_height = m_height + 1&

End Sub

'===========================================================================
' CLongStack.Pop()
'   Pops off the topmost frame
Public Function Pop() As Long

' If m_height > 0, then there is at least one frame
If m_height > 0& Then
  
  ' Reduce the stack height by one
  m_height = m_height - 1&
  
  ' Save the value of the topmost frame to the return value
  Pop = m_stack(m_height)
  
  ' If there are now no frames left, erase the array,
  ' else reduce by one element
  If m_height = 0& Then
    Erase m_stack
  Else
    ReDim Preserve m_stack(m_height - 1&)
  End If
  
End If

End Function

'===========================================================================
' CLongStack.Frame()
'   Returns the value of the current (topmost) frame
Public Function Frame() As Long
Frame = m_stack(m_height - 1&)
End Function
