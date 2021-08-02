Imports System.Runtime.InteropServices

Module Module_SendMessage
    Private Enum GetWindowType As UInteger
        GW_HWNDFIRST = 0
        GW_HWNDLAST = 1
        GW_HWNDNEXT = 2
        GW_HWNDPREV = 3
        GW_OWNER = 4
        GW_CHILD = 5
        GW_ENABLEDPOPUP = 6
    End Enum

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As IntPtr
    End Function
    <DllImport("User32.dll")>
    Private Function FindWindow(ByVal className As String, ByVal caption As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function GetWindow(ByVal hWnd As IntPtr, ByVal uCmd As GetWindowType) As IntPtr
    End Function
    <DllImport("USER32.DLL", EntryPoint:="GetWindowText", SetLastError:=True,
    CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)>
    Private Function GetActiveWindowText(ByVal hWnd As IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
    End Function

    Const WM_SETTEXT As Long = &HC

    ''' <summary>
    ''' Send data to an external application using the SendMessage WINAPI
    ''' </summary>
    ''' <param name="WindowCaption">Target Application Window Text</param>
    ''' <param name="ChildCaption">Target Child Text</param>
    ''' <param name="Message">Payload data</param>
    ''' <returns></returns>
    Public Function SendData(WindowCaption As String, ChildCaption As String, Message As String) As Boolean
        'Find Target Application Handle
        Dim Hwnd As IntPtr = FindWindow(vbNullString, WindowCaption)
        'Get first child handle
        Dim HwndChild As IntPtr = GetWindow(Hwnd, GetWindowType.GW_CHILD)
        'Create Dictionary for storing unique child handles
        Dim ChildHandles As New Dictionary(Of IntPtr, Long)
        Do Until HwndChild.Equals(IntPtr.Zero)
            'Loop through all child handles
            Dim Caption As New System.Text.StringBuilder(256)
            'Declare storage for child caption text
            GetActiveWindowText(HwndChild, Caption, Caption.Capacity)
            'Get child caption text
            If Caption.ToString = ChildCaption Then
                'If target child caption text is found, send the message to it and exit the function
                SendData = SendMessage(HwndChild, WM_SETTEXT, Message.Length, Message)
                Exit Function
            End If
            'Store all unique child handles in the dictionary so we don't infinitely loop
            If Not ChildHandles.ContainsKey(HwndChild) Then
                ChildHandles.Add(HwndChild, 0)
            Else
                Exit Function
            End If
            'Proceed to the next child handle
            HwndChild = GetWindow(Hwnd, GetWindowType.GW_HWNDNEXT)
        Loop
    End Function
End Module