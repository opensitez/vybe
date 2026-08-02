' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_invoice_pdf_generator_mock
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System.Text

Class InvoiceGenerator
    Public Function RenderInvoice(invoiceId As String, amount As Decimal) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("=== INVOICE ===")
        sb.AppendLine("ID: " & invoiceId)
        sb.AppendLine("Amount: $" & amount)
        Return sb.ToString().Trim()
    End Function
End Class

Module Program
    Sub Main()
        Dim gen As New InvoiceGenerator()
        Dim txt = gen.RenderInvoice("INV-500", 150D)
        __Check(CStr(txt.Contains("INV-500")), "True")
    End Sub
End Module
