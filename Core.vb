﻿'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports mshtml
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.InteropServices

'This module contains this program's core procedures.
Public Module CoreModule
   'The API constants and functions used by this module.
   <DllImport("User32.dll", SetLastError:=True)> Private Function EnumChildWindows(ByVal hWndParent As IntPtr, ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function EnumWindows(ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
   End Function
   <DllImport("Oleacc.dll", SetLastError:=True)> Private Function ObjectFromLresult(ByVal lResult As IntPtr, ByRef riid As Guid, ByVal wParam As IntPtr, <MarshalAs(UnmanagedType.Interface)> ByRef ppvObject As mshtml.HTMLDocument) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function RegisterWindowMessageA(ByVal lpString As String) As Integer
   End Function
   <DllImport("user32.dll", SetLastError:=True)> Private Function SendMessageTimeoutA(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByVal flags As Integer, ByVal timeout As Integer, ByRef result As IntPtr) As IntPtr
   End Function

   Private Const SMTO_ABORTIFHUNG As Integer = &H2%

   'The delegates used by this module.
   Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Integer

   'The structures and variables used by this module.

   'This structure defines a HTML document element's attributes
   Public Structure HTMLAttributeStr
      Public Name As String    'Defines an attribute's name.
      Public Value As String   'Defines an attribute's value.
   End Structure

   'This structure defines a HTML document and its elements.
   Public Structure HTMLDocumentStr
      Public Document As mshtml.HTMLDocument       'Defines a document interface.
      Public Elements As List(Of HTMLElementStr)   'Defines a document's elements.
      Public Executable As String                  'Defines the executable displaying a document. 
   End Structure

   'This structure defines a HTML document element.
   Public Structure HTMLElementStr
      Public Attributes As List(Of HTMLAttributeStr)   'Defines an element's attributes.
      Public Name As String                            'Defines an element's name.
   End Structure

   Private ReadOnly DocumentREFIID As New Guid("{626FC520-A41E-11CF-A731-00A0C9082637}")              'Contains the HTML document interface's reference id.
   Private ReadOnly WMHTMLGetObjectMessage As Integer = RegisterWindowMessageA("WM_HTML_GETOBJECT")   'Contains the message used to retrieve a HTML document interface.

   Private HTMLDocuments As List(Of HTMLDocumentStr) = Nothing  'Contains the list of HTML documents and their elements.

   'This procedure checks for HTML document interfaces and add any found to a list.
   Private Sub CheckForDocument(WindowH As IntPtr)
      Try
         Dim Document As mshtml.HTMLDocument = Nothing
         Dim LResult As IntPtr = IntPtr.Zero
         Dim ProcessId As Integer = Nothing

         SendMessageTimeoutA(WindowH, WMHTMLGetObjectMessage, IntPtr.Zero, IntPtr.Zero, SMTO_ABORTIFHUNG, CInt(1000), LResult)
         If Not LResult = IntPtr.Zero Then
            ObjectFromLresult(LResult, DocumentREFIID, IntPtr.Zero, Document)
            If Document IsNot Nothing Then
               GetWindowThreadProcessId(WindowH, ProcessId)
               HTMLDocuments.Add(New HTMLDocumentStr With {.Document = Document, .Elements = GetDocumentElements(Document), .Executable = Process.GetProcessById(ProcessId).MainModule.FileName})
               CheckForFrames(HTMLDocuments.Last)
            End If
         End If
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try
   End Sub

   'This procedure checks for frames in the specified HTML document.
   Private Sub CheckForFrames(DocumentO As HTMLDocumentStr)
      Try
         Dim Frame As mshtml.HTMLDocument = Nothing
         Dim FrameIndex() As Integer = Nothing
         Dim NextFrame As mshtml.HTMLDocument = Nothing
         Dim Parents() As mshtml.HTMLDocument = Nothing
         Dim Level As Integer = 0

         ReDim FrameIndex(0 To Level)
         ReDim Parents(0 To Level)
         Frame = DocumentO.Document
         Do Until (Level = 0) AndAlso (FrameIndex(Level) >= Frame.frames.length)
            Do While FrameIndex(Level) < Frame.frames.length
               NextFrame = DirectCast(DirectCast(Frame.frames.item(CObj(FrameIndex(Level))), mshtml.HTMLWindow2).document, mshtml.HTMLDocument)
               If NextFrame Is Nothing Then Exit Do
               Level += 1
               ReDim Preserve FrameIndex(0 To Level)
               ReDim Preserve Parents(0 To Level)
               Parents(Level) = Frame
               Frame = NextFrame
            Loop

            If NextFrame Is Nothing Then
               HTMLDocuments.Add(New HTMLDocumentStr With {.Document = Frame, .Elements = GetDocumentElements(Frame)})
            Else
               HTMLDocuments.Add(New HTMLDocumentStr With {.Document = Frame, .Elements = GetDocumentElements(Frame)})
               Frame = Parents(Level)
               Level -= 1
               ReDim Preserve FrameIndex(0 To Level)
               ReDim Preserve Parents(0 To Level)
            End If

            If FrameIndex(Level) < Frame.frames.length Then FrameIndex(Level) += 1
         Loop
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try
   End Sub

   'This procedure displays the elements contained by the specified HTML documents.
   Private Sub DisplayDocuments(HTMLDocuments As List(Of HTMLDocumentStr))
      Try
         For Each DocumentO As HTMLDocumentStr In HTMLDocuments
            Console.WriteLine(DocumentO.Executable)
            Console.WriteLine(DocumentO.Document.title)
            For Each ElementO As HTMLElementStr In DocumentO.Elements
               Console.WriteLine($"<{ElementO.Name}>")
               For Each AttributeO As HTMLAttributeStr In ElementO.Attributes
                  Console.WriteLine($"{AttributeO.Name} = {AttributeO.Value}")
               Next AttributeO
            Next ElementO
         Next DocumentO
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try
   End Sub

   'This procedure attempts to enter debug mode and returns the result.
   Private Function EnterDebugMode() As Boolean
      Try
         Process.EnterDebugMode()

         Return True
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try

      Return False
   End Function

   'This procedure returns any elements contained by the specified document.
   Private Function GetDocumentElements(DocumentO As mshtml.HTMLDocument) As List(Of HTMLElementStr)
      Try
         Dim Elements As New List(Of HTMLElementStr)

         With DocumentO.all
            For ItemIndex As Integer = 0 To .length - 1
               GetItemElements(Elements, DirectCast(.item(ItemIndex), IHTMLElement))
            Next ItemIndex
         End With

         Return Elements
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try

      Return Nothing
   End Function

   'This procedure handles any child windows that are found.
   Private Function HandleChildWindow(hWnd As IntPtr, lParam As IntPtr) As Integer
      Try
         CheckForDocument(hWnd)
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try

      Return CInt(True)
   End Function

   'This procedure handles any windows that are found.
   Private Function HandleWindow(hWnd As IntPtr, lParam As IntPtr) As Integer
      Try
         CheckForDocument(hWnd)

         EnumChildWindows(hWnd, AddressOf HandleChildWindow, IntPtr.Zero)
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try

      Return CInt(True)
   End Function

   'This procedue is executed when this program is started.
   Public Sub Main()
      Try
         Dim InDebugMode As New Boolean

         If Environment.GetCommandLineArgs.Last.Trim() = "/?" Then
            Console.WriteLine(ProgramInformation())
            Console.WriteLine()
            Console.WriteLine(My.Application.Info.Description)
         Else
            HTMLDocuments = New List(Of HTMLDocumentStr)

            InDebugMode = EnterDebugMode()
            EnumWindows(AddressOf HandleWindow, IntPtr.Zero)
            If InDebugMode Then Process.LeaveDebugMode()

            If HTMLDocuments.Count = 0 Then
               Console.WriteLine("No HTML documents found.")
            Else
               DisplayDocuments(HTMLDocuments)
            End If
         End If
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try
   End Sub

   'This procedure returns information about this program.
   Private Function ProgramInformation() As String
      Try
         Dim Information As String = Nothing

         With My.Application.Info
            Information = $"{ .Title} v{ .Version} - by: { .CompanyName}"
         End With

         Return Information
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try

      Return Nothing
   End Function
End Module
