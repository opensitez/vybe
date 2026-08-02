' vybe-test: vb/vb_encoding_utf8_ascii_bytes/test_vb_encoding_custom_decoder_fallback_exception
' origin: languages/vb/tests/vb/test_vb_encoding_utf8_ascii_bytes.rs

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

Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim enc As Encoding = Encoding.GetEncoding("utf-8", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
        Dim invalidBytes As Byte() = {&HFE, &HFF} ' Invalid UTF-8 sequence
        Try
            enc.GetString(invalidBytes)
        Catch ex As DecoderFallbackException
            __Check(CStr("DecoderFallbackException Caught"), "DecoderFallbackException Caught")
        End Try
    End Sub
End Module
