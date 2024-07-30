'This module contains code that requires late binding.
Option Strict Off

'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off

Imports mshtml
Imports System
Imports System.Collections.Generic
Imports System.Linq

'This module handles HTML elements.
Public Module ElementsModule
   'This procedure returns the specified item's name and attributes.
   Public Sub GetItemElements(Elements As List(Of HTMLElementStr), Item As IHTMLElement)
      Try
         Dim Attribute As IHTMLDOMAttribute = Nothing
         Dim Value As String = Nothing

         Elements.Add(New HTMLElementStr With {.Attributes = New List(Of HTMLAttributeStr), .Name = Item.tagName})

         If Item.attributes IsNot Nothing Then
            For Nodeindex As Integer = 0 To Item.attributes.length - 1
               Attribute = CType(Item.attributes.item(Nodeindex), IHTMLDOMAttribute)
               Value = Attribute.nodeValue?.ToString()
               If Not Value = Nothing Then
                  Elements.Last().Attributes.Add(New HTMLAttributeStr With {.Name = Attribute.nodeName, .Value = Value})
               End If
            Next Nodeindex
         End If
      Catch ExceptionO As Exception
         Console.Error.WriteLine(ExceptionO.Message)
      End Try
   End Sub
End Module
