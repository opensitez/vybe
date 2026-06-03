use super::helpers::run_vb;

#[test]
fn simple_class_with_fields() {
    let out = run_vb(
        r#"
Module Program
    Class Person
        Public Name As String
        Public Age As Integer
    End Class

    Sub Main()
        Dim p As New Person()
        p.Name = "Alice"
        p.Age = 30
        Console.WriteLine(p.Name)
        Console.WriteLine(p.Age)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alice", "30"]);
}

#[test]
fn class_with_constructor() {
    let out = run_vb(
        r#"
Module Program
    Class Person
        Public Name As String
        Public Age As Integer

        Sub New(n As String, a As Integer)
            Me.Name = n
            Me.Age = a
        End Sub
    End Class

    Sub Main()
        Dim p As New Person("Bob", 25)
        Console.WriteLine(p.Name & " is " & CStr(p.Age))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Bob is 25"]);
}

#[test]
fn class_with_method() {
    let out = run_vb(
        r#"
Module Program
    Class Calculator
        Public Result As Double

        Sub New()
            Me.Result = 0
        End Sub

        Function Add(a As Double, b As Double) As Double
            Add = a + b
        End Function

        Sub AddToResult(value As Double)
            Me.Result = Me.Result + value
        End Sub
    End Class

    Sub Main()
        Dim calc As New Calculator()
        Console.WriteLine(calc.Add(3, 4))
        calc.AddToResult(10)
        calc.AddToResult(20)
        Console.WriteLine(calc.Result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["7", "30"]);
}

#[test]
fn class_with_field_initializer() {
    let out = run_vb(
        r#"
Module Program
    Class Counter
        Public Count As Integer = 0

        Sub Increment()
            Me.Count = Me.Count + 1
        End Sub
    End Class

    Sub Main()
        Dim c As New Counter()
        c.Increment()
        c.Increment()
        c.Increment()
        Console.WriteLine(c.Count)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn multiple_instances() {
    let out = run_vb(
        r#"
Module Program
    Class Dog
        Public Name As String

        Sub New(n As String)
            Me.Name = n
        End Sub

        Function Speak() As String
            Speak = Me.Name & " says Woof!"
        End Function
    End Class

    Sub Main()
        Dim a As New Dog("Rex")
        Dim b As New Dog("Buddy")
        Console.WriteLine(a.Speak())
        Console.WriteLine(b.Speak())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Rex says Woof!", "Buddy says Woof!"]);
}

#[test]
fn class_inheritance() {
    let out = run_vb(
        r#"
Module Program
    Class Animal
        Public Name As String

        Sub New(n As String)
            Me.Name = n
        End Sub

        Function Describe() As String
            Describe = "Animal: " & Me.Name
        End Function
    End Class

    Class Dog
        Inherits Animal

        Sub New(n As String)
            MyBase.New(n)
        End Sub

        Function Bark() As String
            Bark = Me.Name & " barks!"
        End Function
    End Class

    Sub Main()
        Dim d As New Dog("Rex")
        Console.WriteLine(d.Describe())
        Console.WriteLine(d.Bark())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Animal: Rex", "Rex barks!"]);
}

#[test]
fn class_with_property_get_set() {
    let out = run_vb(
        r#"
Module Program
    Class Temperature
        Private _celsius As Double

        Sub New(c As Double)
            _celsius = c
        End Sub

        Property Celsius() As Double
            Get
                Return _celsius
            End Get
            Set(value As Double)
                _celsius = value
            End Set
        End Property

        Property Fahrenheit() As Double
            Get
                Return _celsius * 9 / 5 + 32
            End Get
            Set(value As Double)
                _celsius = (value - 32) * 5 / 9
            End Set
        End Property
    End Class

    Sub Main()
        Dim t As New Temperature(100)
        Console.WriteLine(t.Celsius)
        Console.WriteLine(t.Fahrenheit)
        t.Fahrenheit = 32
        Console.WriteLine(t.Celsius)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100", "212", "0"]);
}

#[test]
fn class_shared_method() {
    let out = run_vb(
        r#"
Module Program
    Class MathHelper
        Shared Function Add(a As Double, b As Double) As Double
            Add = a + b
        End Function

        Shared Function Multiply(a As Double, b As Double) As Double
            Multiply = a * b
        End Function
    End Class

    Sub Main()
        Console.WriteLine(MathHelper.Add(3, 4))
        Console.WriteLine(MathHelper.Multiply(5, 6))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["7", "30"]);
}

#[test]
fn class_tostring_method() {
    let out = run_vb(
        r#"
Module Program
    Class Point
        Public X As Integer
        Public Y As Integer

        Sub New(x As Integer, y As Integer)
            Me.X = x
            Me.Y = y
        End Sub

        Function ToString() As String
            ToString = "(" & CStr(Me.X) & ", " & CStr(Me.Y) & ")"
        End Function
    End Class

    Sub Main()
        Dim p As New Point(10, 20)
        Console.WriteLine(p.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["(10, 20)"]);
}
