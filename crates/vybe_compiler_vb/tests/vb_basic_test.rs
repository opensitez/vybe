use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::{VM, Value};

fn run_vb(source: &str) -> Vec<String> {
    let program = vybe_parser_basic::parse_program(source)
        .unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let output: Rc<RefCell<Vec<String>>> = Rc::new(RefCell::new(Vec::new()));
    let out = output.clone();

    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut vybe_bytecode::HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.borrow_mut().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);

    let chunks = vybe_compiler_vb::Compiler::new().compile(&program)
        .unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));
    output.borrow().clone()
}

// ============================================================
// Basic output
// ============================================================

#[test]
fn hello_world() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine("Hello, World!")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn print_number() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(42)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn print_multiple_values() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine("one")
        Console.WriteLine("two")
        Console.WriteLine("three")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["one", "two", "three"]);
}

// ============================================================
// Variables and assignment
// ============================================================

#[test]
fn dim_and_assign() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Integer = 10
        Dim y As Integer = 20
        Console.WriteLine(x + y)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn string_variable() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim name As String = "Alice"
        Console.WriteLine("Hello " & name)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello Alice"]);
}

#[test]
fn variable_reassign() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Integer = 5
        x = x * 3
        Console.WriteLine(x)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["15"]);
}

// ============================================================
// Arithmetic
// ============================================================

#[test]
fn arithmetic_operations() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(10 + 3)
        Console.WriteLine(10 - 3)
        Console.WriteLine(10 * 3)
        Console.WriteLine(10 / 4)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["13", "7", "30", "2.5"]);
}

#[test]
fn integer_division() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(10 \ 3)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn modulo_operation() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(10 Mod 3)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn exponent_operation() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(2 ^ 10)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["1024"]);
}

// ============================================================
// If/ElseIf/Else
// ============================================================

#[test]
fn if_then() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Integer = 10
        If x > 5 Then
            Console.WriteLine("big")
        End If
    End Sub
End Module
"#);
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_else() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Integer = 3
        If x > 5 Then
            Console.WriteLine("big")
        Else
            Console.WriteLine("small")
        End If
    End Sub
End Module
"#);
    assert_eq!(out, vec!["small"]);
}

#[test]
fn if_elseif_else() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim x As Integer = 5
        If x > 10 Then
            Console.WriteLine("big")
        ElseIf x > 3 Then
            Console.WriteLine("medium")
        Else
            Console.WriteLine("small")
        End If
    End Sub
End Module
"#);
    assert_eq!(out, vec!["medium"]);
}

// ============================================================
// For loop
// ============================================================

#[test]
fn for_loop() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim total As Integer = 0
        For i As Integer = 1 To 5
            total = total + i
        Next
        Console.WriteLine(total)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn for_loop_with_step() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        For i As Integer = 0 To 10 Step 2
            Console.WriteLine(i)
        Next
    End Sub
End Module
"#);
    assert_eq!(out, vec!["0", "2", "4", "6", "8", "10"]);
}

// ============================================================
// While loop
// ============================================================

#[test]
fn while_loop() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim i As Integer = 0
        While i < 5
            Console.WriteLine(i)
            i = i + 1
        End While
    End Sub
End Module
"#);
    assert_eq!(out, vec!["0", "1", "2", "3", "4"]);
}

// ============================================================
// Do Loop
// ============================================================

#[test]
fn do_while_loop() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim i As Integer = 0
        Do While i < 3
            Console.WriteLine(i)
            i = i + 1
        Loop
    End Sub
End Module
"#);
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn do_loop_until() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim i As Integer = 0
        Do
            Console.WriteLine(i)
            i = i + 1
        Loop Until i >= 3
    End Sub
End Module
"#);
    assert_eq!(out, vec!["0", "1", "2"]);
}

// ============================================================
// Sub and Function
// ============================================================

#[test]
fn sub_call() {
    let out = run_vb(r#"
Module Program
    Sub Greet(name As String)
        Console.WriteLine("Hello " & name)
    End Sub

    Sub Main()
        Greet("World")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn function_return() {
    let out = run_vb(r#"
Module Program
    Function Add(a As Integer, b As Integer) As Integer
        Add = a + b
    End Function

    Sub Main()
        Console.WriteLine(Add(3, 4))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn function_with_return_statement() {
    let out = run_vb(r#"
Module Program
    Function Max(a As Integer, b As Integer) As Integer
        If a > b Then
            Return a
        End If
        Return b
    End Function

    Sub Main()
        Console.WriteLine(Max(10, 20))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn recursive_function() {
    let out = run_vb(r#"
Module Program
    Function Factorial(n As Integer) As Integer
        If n <= 1 Then
            Return 1
        End If
        Return n * Factorial(n - 1)
    End Function

    Sub Main()
        Console.WriteLine(Factorial(5))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["120"]);
}

// ============================================================
// Boolean logic
// ============================================================

#[test]
fn boolean_operations() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(True And False)
        Console.WriteLine(True Or False)
        Console.WriteLine(Not True)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["false", "true", "false"]);
}

// ============================================================
// String concatenation
// ============================================================

#[test]
fn string_concat() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Dim a As String = "Hello"
        Dim b As String = " World"
        Console.WriteLine(a & b)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["Hello World"]);
}

// ============================================================
// VB Builtins
// ============================================================

#[test]
fn builtin_ucase_lcase() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(UCase("hello"))
        Console.WriteLine(LCase("WORLD"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["HELLO", "world"]);
}

#[test]
fn builtin_len() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Len("hello"))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn builtin_trim() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Trim("  hello  "))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn builtin_math() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(Abs(-5))
        Console.WriteLine(Math.Floor(3.7))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["5", "3"]);
}

// ============================================================
// Try/Catch
// ============================================================

#[test]
fn try_catch() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Try
            Throw New Exception("oops")
        Catch ex As Exception
            Console.WriteLine("caught")
        End Try
        Console.WriteLine("done")
    End Sub
End Module
"#);
    assert_eq!(out, vec!["caught", "done"]);
}

// ============================================================
// Comparison operators
// ============================================================

#[test]
fn comparison_operators() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Console.WriteLine(5 = 5)
        Console.WriteLine(5 <> 3)
        Console.WriteLine(3 < 5)
        Console.WriteLine(5 > 3)
        Console.WriteLine(3 <= 3)
        Console.WriteLine(5 >= 5)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["true", "true", "true", "true", "true", "true"]);
}

// ============================================================
// Nested function calls
// ============================================================

#[test]
fn nested_calls() {
    let out = run_vb(r#"
Module Program
    Function Double(n As Integer) As Integer
        Double = n * 2
    End Function

    Function AddOne(n As Integer) As Integer
        AddOne = n + 1
    End Function

    Sub Main()
        Console.WriteLine(Double(AddOne(4)))
    End Sub
End Module
"#);
    assert_eq!(out, vec!["10"]);
}

// ============================================================
// Constants
// ============================================================

#[test]
fn constants() {
    let out = run_vb(r#"
Module Program
    Sub Main()
        Const PI As Double = 3.14
        Console.WriteLine(PI)
    End Sub
End Module
"#);
    assert_eq!(out, vec!["3.14"]);
}
