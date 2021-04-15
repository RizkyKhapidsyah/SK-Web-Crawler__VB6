VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W32LIB_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// coded stack routines
Private Stack(50, 9000), sx(50)
Sub LoadDebugWindow()
    For X = 0 To 50
    
    '// no copy if stack is empty
    If GetStackPointer(X) = 0 Then
    Else
    '// copy
    For Y = 0 To GetStackPointer(X)
        debugwin.Win32Stack.Push X, Stack(X, Y)
    Next Y
    End If
    
    Next X
    debugwin.Show
End Sub
Sub WaitForDebugWindow()
    debugwin.CmdFlag = False
    Do
    DoEvents
    Loop Until debugwin.CmdFlag = True
End Sub

Function GetStackPointer(stack_id)
GetStackPointer = sx(stack_id)
End Function
Function Peek(stack_id, stack_index)
Attribute Peek.VB_Description = "Peeks into a stack address. (Can be passed stack pointer!)"
'// return given address in stack
Peek = Stack(stack_id, stack_index)
If Peek = "" Then Peek = False
End Function
Sub CopyStack(stack_id_from, stack_id_to)
Attribute CopyStack.VB_Description = "Copys a stack to another."
'// set counter for destination stack
sx(stack_id_to) = sx(stack_id_from)
For X = 0 To sx(stack_id_from)
    '// copy stack
    Stack(stack_id_to, X) = Stack(stack_id_from, X)
Next X
End Sub
Sub Reset(stack_id)
Attribute Reset.VB_Description = "Erases a stack."
sx(stack_id) = 0 '// reset counter
Stack(stack_id, 0) = 0
End Sub
Sub Push(stack_id, value)
Attribute Push.VB_Description = "pushs a value onto the stack"
Stack(stack_id, sx(stack_id)) = value
 sx(stack_id) = sx(stack_id) + 1
End Sub
Function lne(stack_id)
If sx(stack_id) = 0 Then
lne = True
Else
lne = False
End If
End Function
Function ppeek(stack_id, offset)
Attribute ppeek.VB_Description = "peeks from a offset of current stack pointer."
If sx(stack_id) + offset <= -1 Then
ppeek = False
Exit Function
End If
ppeek = Stack(stack_id, sx(stack_id) + offset)
End Function
Sub poke(stack_id, stack_index, value)
Stack(stack_id, stack_index) = value
End Sub
Function pop(stack_id)
Attribute pop.VB_Description = "pops a value from stack"
If sx(stack_id) - 1 = -1 Then
pop = False
Exit Function
End If
sx(stack_id) = sx(stack_id) - 1
pop = Stack(stack_id, sx(stack_id))
End Function

