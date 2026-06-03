use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Modules, Shared members, scope, OOP patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn module_sub_and_function() {
    let out = run_vb(
        r#"
Module M
    Sub Greet(name As String)
        Console.WriteLine("Hello " & name)
    End Sub
    Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
    Sub Main()
        Greet("World")
        Console.WriteLine(Add(3, 4))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello World", "7"]);
}

#[test]
fn module_level_variables() {
    let out = run_vb(
        r#"
Module M
    Dim counter As Integer = 0
    Sub Increment()
        counter = counter + 1
    End Sub
    Sub Main()
        Increment()
        Increment()
        Increment()
        Console.WriteLine(counter)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn class_shared_method() {
    let out = run_vb(
        r#"
Class MathUtils
    Public Shared Function Square(x As Integer) As Integer
        Return x * x
    End Function
End Class

Module M
    Sub Main()
        Console.WriteLine(MathUtils.Square(5))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn class_with_constructor() {
    let out = run_vb(
        r#"
Class Person
    Public Name As String
    Public Age As Integer
    Public Sub New(n As String, a As Integer)
        Name = n
        Age = a
    End Sub
    Public Function Describe() As String
        Return Name & " is " & CStr(Age)
    End Function
End Class

Module M
    Sub Main()
        Dim p As New Person("Alice", 30)
        Console.WriteLine(p.Describe())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice is 30"]);
}

#[test]
fn class_property_get_set() {
    let out = run_vb(
        r#"
Class Account
    Private _balance As Integer = 0
    Public Property Balance As Integer
        Get
            Return _balance
        End Get
        Set(value As Integer)
            If value >= 0 Then _balance = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim a As New Account()
        a.Balance = 100
        Console.WriteLine(a.Balance)
        a.Balance = -50
        Console.WriteLine(a.Balance)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100", "100"]);
}

#[test]
fn class_inheritance() {
    let out = run_vb(
        r#"
Class Animal
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Function Speak() As String
        Return Name & " makes a sound"
    End Function
End Class

Class Dog
    Inherits Animal
    Public Sub New(n As String)
        MyBase.New(n)
    End Sub
    Public Overrides Function Speak() As String
        Return Name & " says Woof!"
    End Function
End Class

Module M
    Sub Main()
        Dim d As New Dog("Rex")
        Console.WriteLine(d.Speak())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Rex says Woof!"]);
}

#[test]
fn class_me_reference() {
    let out = run_vb(
        r#"
Class Counter
    Public Value As Integer = 0
    Public Sub Increment()
        Me.Value = Me.Value + 1
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        c.Increment()
        c.Increment()
        Console.WriteLine(c.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn multiple_classes() {
    let out = run_vb(
        r#"
Class Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
End Class

Class Segment
    Public Start As Point
    Public Finish As Point
    Public Sub New(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
        Start = New Point(x1, y1)
        Finish = New Point(x2, y2)
    End Sub
    Public Function Length() As Double
        Dim dx As Integer = Finish.X - Start.X
        Dim dy As Integer = Finish.Y - Start.Y
        Return Math.Sqrt(dx * dx + dy * dy)
    End Function
End Class

Module M
    Sub Main()
        Dim l As New Segment(0, 0, 3, 4)
        Console.WriteLine(l.Length())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn recursive_function() {
    let out = run_vb(
        r#"
Module M
    Function Fib(n As Integer) As Integer
        If n <= 1 Then Return n
        Return Fib(n - 1) + Fib(n - 2)
    End Function
    Sub Main()
        Console.WriteLine(Fib(10))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn pass_object_to_sub() {
    let out = run_vb(
        r#"
Class Box
    Public Value As Integer
End Class

Module M
    Sub SetValue(b As Box, v As Integer)
        b.Value = v
    End Sub
    Sub Main()
        Dim b As New Box()
        SetValue(b, 42)
        Console.WriteLine(b.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn function_multiple_return_paths() {
    let out = run_vb(
        r#"
Module M
    Function Classify(x As Integer) As String
        If x > 0 Then Return "positive"
        If x < 0 Then Return "negative"
        Return "zero"
    End Function
    Sub Main()
        Console.WriteLine(Classify(5))
        Console.WriteLine(Classify(-3))
        Console.WriteLine(Classify(0))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["positive", "negative", "zero"]);
}
