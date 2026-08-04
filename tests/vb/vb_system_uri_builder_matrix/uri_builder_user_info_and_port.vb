' vybe-test: vb/vb_system_uri_builder_matrix/uri_builder_user_info_and_port
' origin: languages/vb/tests/vb/test_vb_system_uri_builder_matrix.rs

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

Imports System

Module M
    Sub Main()
        Dim builder As New UriBuilder()
        builder.Scheme = "https"
        builder.Host = "example.com"
        builder.UserName = "alice"
        builder.Password = "secret"
        builder.Port = 8443

        Dim built As Uri = builder.Uri
        __P(CStr(built.UserInfo))
        __P(CStr(built.Authority))
        __P(CStr(built.Port))
        __Check("alice:secret
alice:secret@example.com:8443
8443")
    End Sub
End Module
