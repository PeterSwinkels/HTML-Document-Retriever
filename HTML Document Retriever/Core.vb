'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices

'This module contains this program's core procedures.
Public Module CoreModule
   'The API constants and functions used by this module.
   <DllImport("User32.dll", SetLastError:=True)> Private Function EnumChildWindows(ByVal hWndParent As Integer, ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As Integer) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function EnumWindows(ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As Integer) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function GetWindowThreadProcessId(ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
   End Function
   <DllImport("Oleacc.dll", SetLastError:=True)> Private Function ObjectFromLresult(ByVal lResult As Integer, ByRef riid As Guid, ByVal wParam As Integer, <MarshalAs(UnmanagedType.Interface)> ByRef ppvObject As mshtml.HTMLDocument) As Integer
   End Function
   <DllImport("User32.dll", SetLastError:=True)> Private Function RegisterWindowMessageA(ByVal lpString As String) As Integer
   End Function
   <DllImport("user32.dll", SetLastError:=True)> Private Function SendMessageTimeoutA(ByVal hWnd As Integer, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer, ByVal flags As Integer, ByVal timeout As Integer, ByRef result As Integer) As IntPtr
   End Function

   Private Const SMTO_ABORTIFHUNG As Integer = &H2%

   'The delegates used by this module.
   Private Delegate Function EnumWindowsProc(ByVal hWnd As Integer, ByVal lParam As Integer) As Integer

   'The structures and variables used by this module.

   'This structure defines a HTML document element's attributes
   Private Structure HTMLAttributeStr
      Public Name As String    'Defines an attribute's name.
      Public Value As String   'Defines an attribute's value.
   End Structure

   'This structure defines a HTML document and its elements.
   Private Structure HTMLDocumentStr
      Public Document As mshtml.HTMLDocument       'Defines a document interface.
      Public Elements As List(Of HTMLElementStr)   'Defines a document's elements.
      Public Executable As String                  'Defines the executable displaying a document. 
   End Structure

   'This structure defines a HTML document element.
   Private Structure HTMLElementStr
      Public Attributes As List(Of HTMLAttributeStr)   'Defines an element's attributes.
      Public Name As String                            'Defines an element's name.
   End Structure

   Private ReadOnly DocumentREFIID As New Guid("{626FC520-A41E-11CF-A731-00A0C9082637}")              'Contains the HTML document interface's reference id.
   Private ReadOnly WMHTMLGetObjectMessage As Integer = RegisterWindowMessageA("WM_HTML_GETOBJECT")   'Contains the message used to retrieve a HTML document interface.

   Private HTMLDocuments As List(Of HTMLDocumentStr) = Nothing  'Contains the list of HTML documents and their elements.

   'This procedure checks for HTML document interfaces and add any found to a list.
   Private Sub CheckForDocument(WindowH As Integer)
      Try
         Dim Document As mshtml.HTMLDocument = Nothing
         Dim LResult As Integer = Nothing
         Dim ProcessId As Integer = Nothing

         SendMessageTimeoutA(WindowH, WMHTMLGetObjectMessage, CInt(0), CInt(0), SMTO_ABORTIFHUNG, CInt(1000), LResult)
         If Not LResult = 0 Then
            ObjectFromLresult(LResult, DocumentREFIID, CInt(0), Document)
            If Document IsNot Nothing Then
               GetWindowThreadProcessId(WindowH, ProcessId)
               HTMLDocuments.Add(New HTMLDocumentStr With {.Document = Document, .Elements = GetElements(Document), .Executable = Process.GetProcessById(ProcessId).MainModule.FileName})
            End If
         End If
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try
   End Sub

   'This procedure returns any elements contained by the specified document.
   Private Function GetElements(DocumentO As mshtml.HTMLDocument) As List(Of HTMLElementStr)
      Dim Elements As New List(Of HTMLElementStr)

      Try
         Dim Item As Object = Nothing
         Dim Value As String = Nothing

         With DocumentO.all
            For ItemIndex As Integer = 0 To .length - 1
               Item = .item(ItemIndex)
               Elements.Add(New HTMLElementStr With {.Attributes = New List(Of HTMLAttributeStr), .Name = Item.tagName})
               If Item.Attributes IsNot Nothing Then
                  For Nodeindex As Integer = 0 To Item.Attributes.Length - 1
                     Value = Item.Attributes(Nodeindex).nodeValue?.ToString()
                     If Not Value = Nothing Then
                        Elements.Last().Attributes.Add(New HTMLAttributeStr With {.Name = Item.Attributes(Nodeindex).nodeName, .Value = Value})
                     End If
                  Next NodeIndex
               End If
            Next ItemIndex
         End With
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try

      Return Elements
   End Function

   'This procedure handles any child windows that are found.
   Private Function HandleChildWindow(hWnd As Integer, lParam As Integer) As Integer
      Try
         CheckForDocument(hWnd)
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try

      Return CInt(True)
   End Function

   'This procedure handles any windows that are found.
   Private Function HandleWindow(hWnd As Integer, lParam As Integer) As Integer
      Try
         CheckForDocument(hWnd)

         EnumChildWindows(hWnd, AddressOf HandleChildWindow, Nothing)
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try

      Return CInt(True)
   End Function

   'This procedue is executed when this program is started.
   Public Sub Main()
      Try
         HTMLDocuments = New List(Of HTMLDocumentStr)

         Process.EnterDebugMode()
         EnumWindows(AddressOf HandleWindow, CInt(0))
         Process.LeaveDebugMode()

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
         Console.WriteLine(ExceptionO.Message)
      End Try
   End Sub
End Module
