'This module contains code that requires late binding.
Option Strict Off

'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Linq

'This module handles HTML elements.
Public Module ElementsModule
   'This procedure returns the specified item's name and attributes.
   Public Sub GetItemElements(Elements As List(Of HTMLElementStr), Item As Object)
      Try
         Dim Value As String = Nothing

         Elements.Add(New HTMLElementStr With {.Attributes = New List(Of HTMLAttributeStr), .Name = Item.tagName})
         If Item.Attributes IsNot Nothing Then
            For Nodeindex As Integer = 0 To Item.Attributes.Length - 1
               Value = Item.Attributes(Nodeindex).nodeValue?.ToString()
               If Not Value = Nothing Then
                  Elements.Last().Attributes.Add(New HTMLAttributeStr With {.Name = Item.Attributes(Nodeindex).nodeName, .Value = Value})
               End If
            Next Nodeindex
         End If
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try
   End Sub
End Module
