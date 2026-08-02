' vybe-test: vb/vb_advanced_linq_xml/xml_axis_extension_method
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

Imports System.Xml.Linq
Imports System.Runtime.CompilerServices

Module Extensions
    <Extension()>
    Public Function GetNames(elements As IEnumerable(Of XElement)) As IEnumerable(Of String)
        Return elements.Select(Function(e) e.Name.LocalName)
    End Function
End Module

Module M
    Sub Main()
        Dim xml = <Root><A/><B/></Root>
        Dim names = xml.Elements().GetNames()
        
        For Each name In names
            Console.WriteLine(name)
        Next
    End Sub
End Module
