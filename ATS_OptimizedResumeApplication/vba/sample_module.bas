Attribute VB_Name = "SampleModule"
Option Explicit

' Sample module retained for compatibility; demonstrates calling app entry points.

' PUBLIC_INTERFACE
Sub Hello()
    MsgBox AppInfo(), vbInformation, "Hello"
End Sub

' PUBLIC_INTERFACE
Sub InitializeAppDemo()
    EntryPoint.App_Initialize
End Sub
