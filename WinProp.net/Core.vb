﻿'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Convert
Imports System.Diagnostics
Imports System.Environment
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal
Imports System.Text
Imports System.Threading.Thread

'This module contains this program's core procedures.
Public Module CoreModule
   'The Microsoft Windows API constants, delegates, and functions used by this module.
   Private Const EM_GETPASSWORDCHAR As Integer = &HD2%
   Private Const EM_SETPASSWORDCHAR As Integer = &HCC%
   Private Const ES_PASSWORD As Integer = &H20%
   Private Const GWL_STYLE As Integer = &HFFFFFFF0%
   Private Const WM_GETTEXT As Integer = &HD%
   Private Const WM_GETTEXTLENGTH As Integer = &HE%

   <DllImport("User32.dll", SetLastError:=True)> Private Function EnumChildWindows(ByVal hWndParent As IntPtr, ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function EnumPropsExA(ByVal hwnd As IntPtr, ByVal lpEnumFunc As PropEnumProcEx, ByVal lParam As IntPtr) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function EnumWindows(ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function GetClassNameW(ByVal hWnd As IntPtr, ByVal lpClassName As IntPtr, ByVal nMaxCount As Integer) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function GetParent(ByVal hwnd As IntPtr) As IntPtr
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function GetWindowLongA(ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function PostMessageA(ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function SendMessageW(ByVal hwnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function SetPropA(ByVal hwnd As IntPtr, ByVal lpString As String, ByVal hData As IntPtr) As Integer
   End Function

   Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Integer
   Private Delegate Function PropEnumProcEx(ByVal hwnd As IntPtr, ByVal lpszString As IntPtr, ByVal hData As IntPtr, ByVal dwData As Integer) As Integer

   'This structure defines a window's property.
   Private Structure WindowPropertyStr
      Public Name As String    'Defines a property's name.
      Public Value As String   'Defines a property's value.
   End Structure

   Private WindowProperties As New List(Of WindowPropertyStr)   'Contains the list of all properties contained by all active windows.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         EnumWindows(AddressOf HandleWindow, IntPtr.Zero)
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure displays any errors that occur.
   Private Sub DisplayError(ExceptionO As Exception)
      Try
         Console.ForegroundColor = ConsoleColor.Red
         Console.Error.WriteLine($"{NewLine}ERROR: {ExceptionO.Message}{NewLine}")
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure displays the specified window's information.
   Private Sub DisplayWindowInformation(WindowH As IntPtr)
      Try
         Console.ForegroundColor = ConsoleColor.Yellow
         Console.WriteLine(GetWindowProcessPath(WindowH))
         Console.ForegroundColor = ConsoleColor.Green
         Console.WriteLine(GetWindowPath(WindowH))
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the specified window properties.
   Private Sub DisplayWindowProperties(Properties As List(Of WindowPropertyStr))
      Try
         Console.ForegroundColor = ConsoleColor.Cyan
         WindowProperties.ForEach(Sub([Property]) Console.WriteLine($"""{[Property].Name}"" = ""{[Property].Value}"""))
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure converts non-displayable or all (if specified) characters in the specified text to escape sequences.
   Public Function Escape(ToEscape As String, Optional EscapeCharacter As Char = "/"c) As String
      Try
         Dim Character As New Char
         Dim Escaped As New StringBuilder
         Dim Index As Integer = 0
         Dim Text As String = If(ToEscape Is Nothing, "", ToEscape)

         With Escaped
            Do Until Index >= Text.Length
               Character = Text.Chars(Index)

               If Character = EscapeCharacter Then
                  .Append(New String(EscapeCharacter, 2))
               ElseIf (Character = ControlChars.Tab OrElse (Character >= " "c AndAlso Character <= "~"c)) Then
                  .Append(Character)
               Else
                  .Append($"{EscapeCharacter}{ToInt32(Character):X4}")
               End If
               Index += 1
            Loop
         End With

         Return Escaped.ToString()
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified window's class.
   Private Function GetWindowClass(WindowH As IntPtr) As String
      Try
         Dim Buffer As IntPtr = AllocHGlobal(UShort.MaxValue)
         Dim Length As Integer = CInt(GetClassNameW(WindowH, Buffer, UShort.MaxValue))
         Dim WindowClass As String = If(Length > 0, PtrToStringUni(Buffer).Substring(0, Length), Nothing)

         FreeHGlobal(Buffer)

         Return WindowClass
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified window's class and text preceded by any parent windows' information.
   Private Function GetWindowPath(WindowHandle As IntPtr) As String
      Try
         Dim WindowPath As New StringBuilder

         Do Until WindowHandle = IntPtr.Zero
            If WindowPath.Length > 0 Then WindowPath.Append("\")
            WindowPath.Append($"""{Escape(GetWindowText(WindowHandle)).Trim()}""(""{Escape(GetWindowClass(WindowHandle)).Trim()}"")")
            WindowHandle = GetParent(WindowHandle)
         Loop

         Return WindowPath.ToString()
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified window's process path.
   Private Function GetWindowProcessPath(WindowH As IntPtr) As String
      Try
         Dim ProcessId As New Integer

         GetWindowThreadProcessId(WindowH, ProcessId)

         Return Process.GetProcessById(ProcessId).MainModule.FileName
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the specified window's text.
   Private Function GetWindowText(WindowH As IntPtr) As String
      Try
         Dim Buffer As New IntPtr
         Dim Length As New Integer
         Dim PasswordCharacter As New Integer
         Dim WindowText As String = Nothing

         If WindowHasStyle(WindowH, ES_PASSWORD) Then
            PasswordCharacter = CInt(SendMessageW(WindowH, EM_GETPASSWORDCHAR, IntPtr.Zero, IntPtr.Zero))
            If Not PasswordCharacter = Nothing Then
               PostMessageA(WindowH, EM_SETPASSWORDCHAR, IntPtr.Zero, IntPtr.Zero)
               Sleep(1000)
            End If
         End If

         Length = CInt(SendMessageW(WindowH, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero)) + 1
         Buffer = AllocHGlobal(UShort.MaxValue)
         Length = CInt(SendMessageW(WindowH, WM_GETTEXT, CType(Length, IntPtr), Buffer))
         WindowText = PtrToStringUni(Buffer)
         FreeHGlobal(Buffer)

         WindowText = If(Length <= WindowText.Length, WindowText.Substring(0, Length), Nothing)

         If Not PasswordCharacter = Nothing Then PostMessageA(WindowH, EM_SETPASSWORDCHAR, CType(PasswordCharacter, IntPtr), IntPtr.Zero)

         Return WindowText
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure handles the specified child window.
   Private Function HandleChildWindow(hWnd As IntPtr, lParam As IntPtr) As Integer
      Try
         WindowProperties.Clear()
         EnumPropsExA(hWnd, AddressOf HandleWindowProperty, IntPtr.Zero)
         If WindowProperties.Count > 0 Then
            Console.WriteLine()
            DisplayWindowInformation(hWnd)
            DisplayWindowProperties(WindowProperties)
         End If

         Return CInt(True)
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return CInt(True)
   End Function

   'This procedure handles the specified window.
   Private Function HandleWindow(hWnd As IntPtr, lParam As IntPtr) As Integer
      Try
         WindowProperties.Clear()

         EnumPropsExA(hWnd, AddressOf HandleWindowProperty, IntPtr.Zero)
         If WindowProperties.Count > 0 Then
            Console.WriteLine()
            DisplayWindowInformation(hWnd)
            DisplayWindowProperties(WindowProperties)
         End If

         EnumChildWindows(hWnd, AddressOf HandleChildWindow, IntPtr.Zero)

         Return CInt(True)
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return CInt(True)
   End Function

   'This procedure handles the specified window property.
   Private Function HandleWindowProperty(ByVal hwnd As IntPtr, ByVal lpszString As IntPtr, ByVal hData As IntPtr, ByVal dwData As Integer) As Integer
      Try
         WindowProperties.Add(New WindowPropertyStr With {.Name = PtrToStringAnsi(lpszString), .Value = Escape(PtrToStringAnsi(hData))})

         Return CInt(True)
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return CInt(True)
   End Function

   'This procedure returns the checks whether a window has the specified style and returns the result.
   Private Function WindowHasStyle(WindowH As IntPtr, Style As Integer) As Boolean
      Try
         Return (CInt(GetWindowLongA(WindowH, GWL_STYLE)) And Style) = Style
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
