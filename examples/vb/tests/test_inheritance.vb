Module MainModule
    Sub Main()
        ' ── Basic Inheritance ─────────────────────────────────
        Dim d As New Derived
        d.BaseField = 10
        d.DerivedField = 20
        Console.WriteLine("BaseField: " & d.BaseField) ' Should be 10
        Console.WriteLine("DerivedField: " & d.DerivedField) ' Should be 20
        d.BaseMethod() ' Should print "Base Method"
        d.DerivedMethod() ' Should print "Derived Method"

        ' ── Partial Classes ───────────────────────────────────
        Dim p As New PartialClass
        p.Part1Method()
        p.Part2Method()

        ' ── MyBase ────────────────────────────────────────────
        Console.WriteLine("--- MyBase Tests ---")
        Dim c As New Child
        c.Greet()
        ' Should print:
        '   Parent says: Hello
        '   Child says: Hi

        ' ── Constructor Chaining ──────────────────────────────
        Console.WriteLine("--- Constructor Chaining ---")
        Dim cc As New ChildWithCtor
        Console.WriteLine("Message: " & cc.Message)
        ' If chaining works, Message should be "Base initialized"

        ' ── TypeOf...Is with Hierarchy ────────────────────────
        Console.WriteLine("--- TypeOf...Is Tests ---")
        Dim obj As New GrandChild
        Console.WriteLine("GrandChild Is GrandChild: " & (TypeOf obj Is GrandChild))  ' True
        Console.WriteLine("GrandChild Is Child: " & (TypeOf obj Is Child))             ' True
        Console.WriteLine("GrandChild Is Parent: " & (TypeOf obj Is Parent))           ' True

        ' ── Method Overriding ─────────────────────────────────
        Console.WriteLine("--- Override Tests ---")
        Dim ov As New OverrideDerived
        ov.Speak()
        ' Should print: "Override: I am the derived version"

        ' ── Multi-level Inheritance ───────────────────────────
        Console.WriteLine("--- Multi-Level Inheritance ---")
        Dim gc As New GrandChild
        gc.ParentDo()    ' "Parent does"
        gc.ChildDo()     ' "Child does"
        gc.GrandChildDo() ' "GrandChild does"

        ' ── MyBase.New() with explicit call ───────────────────
        Console.WriteLine("--- Explicit MyBase.New ---")
        Dim ec As New ExplicitChild
        Console.WriteLine("Base init: " & ec.BaseInit)
        Console.WriteLine("Child init: " & ec.ChildInit)

        Console.WriteLine("All inheritance tests passed!")
    End Sub
End Module

' ── Basic classes ─────────────────────────────────────────────
Class Base
    Public BaseField As Integer
    Public Sub BaseMethod()
        Console.WriteLine("Base Method")
    End Sub
End Class

Class Derived
    Inherits Base
    Public DerivedField As Integer
    Public Sub DerivedMethod()
        Console.WriteLine("Derived Method")
    End Sub
End Class

' ── Partial classes ───────────────────────────────────────────
Partial Public Class PartialClass
    Public Sub Part1Method()
        Console.WriteLine("Part 1")
    End Sub
End Class

Partial Public Class PartialClass
    Public Sub Part2Method()
        Console.WriteLine("Part 2")
    End Sub
End Class

' ── MyBase test ───────────────────────────────────────────────
Class Parent
    Public Overridable Sub Greet()
        Console.WriteLine("Parent says: Hello")
    End Sub

    Public Sub ParentDo()
        Console.WriteLine("Parent does")
    End Sub
End Class

Class Child
    Inherits Parent

    Public Overrides Sub Greet()
        MyBase.Greet()
        Console.WriteLine("Child says: Hi")
    End Sub

    Public Sub ChildDo()
        Console.WriteLine("Child does")
    End Sub
End Class

Class GrandChild
    Inherits Child

    Public Sub GrandChildDo()
        Console.WriteLine("GrandChild does")
    End Sub
End Class

' ── Constructor chaining ──────────────────────────────────────
Class BaseWithCtor
    Public Message As String

    Public Sub New()
        Message = "Base initialized"
    End Sub
End Class

Class ChildWithCtor
    Inherits BaseWithCtor

    ' No explicit Sub New — base Sub New should be auto-called
End Class

' ── Method override ───────────────────────────────────────────
Class OverrideBase
    Public Overridable Sub Speak()
        Console.WriteLine("I am the base version")
    End Sub
End Class

Class OverrideDerived
    Inherits OverrideBase

    Public Overrides Sub Speak()
        Console.WriteLine("Override: I am the derived version")
    End Sub
End Class

' ── Explicit MyBase.New() ─────────────────────────────────────
Class ExplicitBase
    Public BaseInit As Boolean

    Public Sub New()
        BaseInit = True
    End Sub
End Class

Class ExplicitChild
    Inherits ExplicitBase
    Public ChildInit As Boolean

    Public Sub New()
        MyBase.New()
        ChildInit = True
    End Sub
End Class
