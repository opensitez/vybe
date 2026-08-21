' vybe-test: vb/vb_xml_linq_transformation_pipeline/test_vb_xml_domain_objects_to_xml_transformation
' origin: languages/vb/tests/vb/test_vb_xml_linq_transformation_pipeline.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Imports System.Collections.Generic
Imports System.Linq
Imports System.Xml.Linq
Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module


Class Item
    Public Property Code As String
    Public Property Qty As Integer
End Class

Module Program
    Sub Main()
        Dim items As New List(Of Item) From {
            New Item With {.Code = "I1", .Qty = 10},
            New Item With {.Code = "I2", .Qty = 20}
        }

        Dim root As New XElement("Inventory",
            From i In items Select New XElement("Item", New XAttribute("Code", i.Code), i.Qty)
        )

        __P(CStr(root.Elements("Item").Count() & "|" & root.ToString().Contains("Code=""I1""")))
        __Check("2|True")
    End Sub
End Module
