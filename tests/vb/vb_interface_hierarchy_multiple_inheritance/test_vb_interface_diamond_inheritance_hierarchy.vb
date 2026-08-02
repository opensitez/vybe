' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_diamond_inheritance_hierarchy
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IBase
    Sub BaseMethod()
End Interface

Interface ILeft
    Inherits IBase
    Sub LeftMethod()
End Interface

Interface IRight
    Inherits IBase
    Sub RightMethod()
End Interface

Interface IDiamond
    Inherits ILeft, IRight
    Sub DiamondMethod()
End Interface

Class DiamondImpl
    Implements IDiamond
    Public Sub BaseMethod() Implements IBase.BaseMethod
        __Check(CStr("BaseMethod"), "BaseMethod")
    End Sub
    Public Sub LeftMethod() Implements ILeft.LeftMethod
        __Check(CStr("LeftMethod"), "LeftMethod")
    End Sub
    Public Sub RightMethod() Implements IRight.RightMethod
        __Check(CStr("RightMethod"), "RightMethod")
    End Sub
    Public Sub DiamondMethod() Implements IDiamond.DiamondMethod
        __Check(CStr("DiamondMethod"), "DiamondMethod")
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As IDiamond = New DiamondImpl()
        d.BaseMethod()
        d.LeftMethod()
        d.RightMethod()
        d.DiamondMethod()
    End Sub
End Module
