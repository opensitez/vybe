//! Local functions: declaration, recursion, closure capture, static local.
use super::helpers::run_csharp;

#[test]
fn local_function_declared_and_called_within_method() {
    assert_eq!(
        run_csharp(r#"int Square(int n){
    int Sq(int x)=>x*x;
    return Sq(n);
}
Console.WriteLine(Square(5));"#),
        &["25"]
    );
}

#[test]
fn recursive_local_function_computes_fibonacci() {
    assert_eq!(
        run_csharp(r#"int Fib(int n){
    int F(int k)=>k<=1?k:F(k-1)+F(k-2);
    return F(n);
}
Console.WriteLine(Fib(7));"#),
        &["13"]
    );
}

#[test]
fn local_function_captures_outer_variable() {
    assert_eq!(
        run_csharp(r#"int multiplier=3;
int Mul(int n){
    int Scaled(int x)=>x*multiplier;
    return Scaled(n);
}
Console.WriteLine(Mul(7));"#),
        &["21"]
    );
}

#[test]
fn static_local_function_cannot_capture_outer_variable() {
    assert_eq!(
        run_csharp(r#"static int Pure(int a,int b){
    static int Add(int x,int y)=>x+y;
    return Add(a,b);
}
Console.WriteLine(Pure(4,5));"#),
        &["9"]
    );
}

#[test]
fn local_function_used_before_its_declaration() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(Double(5));
int Double(int x)=>x*2;"#),
        &["10"]
    );
}

#[test]
fn local_function_returns_func_delegate() {
    assert_eq!(
        run_csharp(r#"System.Func<int,int> MakeAdder(int n){
    int Add(int x)=>x+n;
    return Add;
}
var add10=MakeAdder(10);
Console.WriteLine(add10(5));"#),
        &["15"]
    );
}
