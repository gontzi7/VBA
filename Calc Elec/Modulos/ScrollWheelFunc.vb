Option Explicit

'This will compile in 32 bit Excel only


#If Win64 Then
Public Declare PtrSafe Function FindWindowA& Lib "user32" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare PtrSafe Function GetWindowLongA& Lib "user32" (ByVal hwnd&, ByVal nIndex&)

Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                       (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)

Private Declare PtrSafe Function SetWindowsHookEx Lib _
                                  "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, _
                                                                      ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
                                              ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#Else
Public Declare Function FindWindowA& Lib "user32" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function GetWindowLongA& Lib "user32" (ByVal hwnd&, ByVal nIndex&)
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                       (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Declare Function SetWindowsHookEx Lib _
                                  "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
                                                                      ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
                                              ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type MSLLHOOKSTRUCT  'Will Hold the lParam struct Data
  pt As POINTAPI
  mouseData As Long  ' Holds Forward\Backward flag
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private Const HC_ACTION = 0
Private Const WH_MOUSE_LL = 14
Private Const WM_MOUSEWHEEL = &H20A
Private Const GWL_HINSTANCE = (-6)

Public Const nMyControlTypeNONE = 0
Public Const nMyControlTypeUSERFORM = 1
Public Const nMyControlTypeFRAME = 2
Public Const nMyControlTypeCOMBOBOX = 3
Public Const nMyControlTypeLISTBOX = 4

Private mInScroll As Boolean
Private mLastTick As Double

Private lastWheelTick As Long        ' timestamp of last processed wheel event (ms)
Private wheelProcessing As Boolean   ' guard to prevent reentrancy
Private Const WHEEL_DEBOUNCE_MS As Long = 25  ' threshold in ms (tuneable)

Private hhkLowLevelMouse As Long
Private udtlParamStuct As MSLLHOOKSTRUCT

Public myGblUserForm As UserForm
Public myGblControlObject As Object
'Public iGblControlType As Integer

Public myGblUserFormControl As Object
Private Function FindControlType(myObject As Object)
    FindControlType = nMyControlTypeNONE
    If TypeOf myObject Is MSForms.ComboBox Then
        FindControlType = nMyControlTypeCOMBOBOX
    Else
        If TypeOf myObject Is MSForms.Frame Then
            FindControlType = nMyControlTypeFRAME
        Else
            If TypeOf myObject Is MSForms.UserForm Then
                FindControlType = nMyControlTypeUSERFORM
            Else
                If TypeOf myObject Is MSForms.ListBox Then
                    FindControlType = nMyControlTypeLISTBOX
                End If
            End If
        End If
    End If
End Function

#If Win64 Then
Function GetHookStruct(ByVal lParam As LongPtr) As MSLLHOOKSTRUCT
#Else
Function GetHookStruct(ByVal lParam As Long) As MSLLHOOKSTRUCT
#End If
' VarPtr returns address; LenB returns size in bytes.
  CopyMemory VarPtr(udtlParamStuct), lParam, LenB(udtlParamStuct)
  GetHookStruct = udtlParamStuct
End Function
#If Win64 Then
Function LowLevelMouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As LongPtr
#Else
Function LowLevelMouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
  'Avoid XL crashing if RunTime error occurs due to Mouse fast movement
  
  Dim iDirection As Long

  On Error GoTo HookError
  '    \\ Unhook & get out in case the application is deactivated
  If GetForegroundWindow <> FindWindowA("ThunderDFrame", myGblUserForm.Caption) Then
    UnHook_Mouse
    Exit Function
  End If
  If (nCode = HC_ACTION) Then
    If wParam = WM_MOUSEWHEEL Then
    
      iDirection = GetHookStruct(lParam).mouseData
      ProcessMouseWheelMovement iDirection
    
      '\\ Don't process Default WM_MOUSEWHEEL Window message
      LowLevelMouseProc = True
    End If
    
    Exit Function
  End If
  LowLevelMouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    Exit Function
HookError:
    ' If anything odd happens, try to gracefully unhook and avoid crashing Excel
    On Error Resume Next
    wheelProcessing = False
    UnHook_Mouse
    LowLevelMouseProc = 1
End Function

Public Sub Hook_Mouse(myObject As Object)
    Dim i As Integer
    Set myGblControlObject = myObject
    On Error GoTo GoHook
    Do While Not myObject Is Nothing
        Set myObject = myObject.Parent
    Loop
GoHook:
    Set myGblUserForm = myObject
    Set myObject = Nothing
    sHook_Mouse
End Sub

Private Sub sHook_Mouse()
' Statement to maintain the handle of the hook if clicking outside of the control.
' There isn't a Hinstance for Application, so used GetWindowLong to get handle.
  If hhkLowLevelMouse < 1 Then
    hhkLowLevelMouse = SetWindowsHookEx(WH_MOUSE_LL, AddressOf LowLevelMouseProc, _
      GetWindowLongA(FindWindowA("ThunderDFrame", myGblUserForm.Caption), GWL_HINSTANCE), 0)
  End If
End Sub

Public Sub UnHook_Mouse()
  If hhkLowLevelMouse <> 0 Then
    
    UnhookWindowsHookEx hhkLowLevelMouse
    hhkLowLevelMouse = 0
  End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UserForm MouseWheel Code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ProcessMouseWheelMovement(ByVal iDirection As Long)
    On Error GoTo ProcError

    Const WHEEL_DEBOUNCE_MS As Double = 0.01 ' 10 ms
    
    ' === Reentrancy guard ===
    If mInScroll Then Exit Sub
    mInScroll = True
    
    ' === Debounce ===
    If Timer - mLastTick < WHEEL_DEBOUNCE_MS Then
        mInScroll = False
        Exit Sub
    End If
    mLastTick = Timer

    If myGblControlObject Is Nothing Then Exit Sub
    If myGblUserForm Is Nothing Then Exit Sub

    Dim ctrlType As Long
    ctrlType = FindControlType(myGblControlObject)

    Select Case ctrlType
        Case nMyControlTypeUSERFORM, nMyControlTypeFRAME
            Dim i As Long, mult As Long
            mult = IIf(ctrlType = nMyControlTypeUSERFORM, 3, 3)
            If iDirection > 0 Then
                For i = 1 To mult
                    On Error Resume Next
                    myGblControlObject.Scroll fmScrollActionNoChange, fmScrollActionLineUp
                    On Error GoTo ProcError
                Next i
            Else
                For i = 1 To mult
                    On Error Resume Next
                    myGblControlObject.Scroll fmScrollActionNoChange, fmScrollActionLineDown
                    On Error GoTo ProcError
                Next i
            End If

        Case nMyControlTypeCOMBOBOX, nMyControlTypeLISTBOX
            With myGblControlObject
                Dim newTop As Long
                newTop = .TopIndex + IIf(iDirection > 0, -2, 2)
                If newTop < 0 Then newTop = 0
                If .ListCount > 0 Then
                    If newTop > .ListCount - 1 Then newTop = .ListCount - 1
                Else
                    newTop = 0
                End If
                ' Protect assignment
                On Error Resume Next
                .TopIndex = newTop
                On Error GoTo ProcError
            End With

        Case Else
            ' Nothing to do
    End Select

    mInScroll = False

    Exit Sub

ProcError:
    ' On any processing error, stop processing and unhook to prevent repeated crashes
    On Error Resume Next
    wheelProcessing = False
    UnHook_Mouse
End Sub