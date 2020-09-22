VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================================================================
' CStringStack
'---------------------------------------------------------------------------
'
' Name:           CStringStack
' Author:         Daniel Keep
' Contact:        shortcircuitfky@hotmail.com
'                 When contacting me, please include [VB] in the subject
'                 line, or I might not read your email.
' Version:        1.0
' Last Modified:  6 June 2002
' Description:    Implements a light-weight string stack class.
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
Private m_stack() As String

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
' CStringStack.Push()
'   Pushes a new frame onto the stack
Public Sub Push(ByRef Frame$)

' Resize the stack array to accomodate the new frame
ReDim Preserve m_stack(m_height)

' Add the new frame
m_stack(m_height) = Frame
m_height = m_height + 1&

End Sub

'===========================================================================
' CStringStack.Pop()
'   Pops off the topmost frame
Public Function Pop() As String

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
' CStringStack.Frame()
'   Returns the value of the current (topmost) frame
Public Function Frame() As String
Frame = m_stack(m_height - 1&)
End Function

'===========================================================================
' CStringStack.DumpStack()
'   This is a hang-over from when I wrote my original VBStack class for
'   keeping a call stack for a scripting runtime.  This just dumps all the
'   frames of the stack to the return value, seperated by [Delimiter]
Public Function DumpStack(Optional ByRef Delimiter As String = "/") As String

' Create counter (i&), and temp variable (stack$)
Dim stack$, i&

' If there is at least one frame, save the first one
If m_height > 0& Then stack = m_stack(0&)

' Now, starting with the second frame, start concatenating
For i = 1& To m_height - 1&
  stack = stack & Delimiter & m_stack(i)
Next i

' Return the stack dump
DumpStack = stack

' Erase the temp variable
stack = ""

End Function
