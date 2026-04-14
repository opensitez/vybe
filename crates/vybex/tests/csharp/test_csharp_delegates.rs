use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Delegates, events, Action/Func, callbacks
// ═══════════════════════════════════════════════════════════

#[test]
fn action_delegate() {
    let out = run_csharp(r#"
Action<string> greet = name => Console.WriteLine("Hello " + name);
greet("World");
greet("Alice");
"#);
    assert_eq!(out, vec!["Hello World", "Hello Alice"]);
}

#[test]
fn func_delegate() {
    let out = run_csharp(r#"
Func<int, int> square = x => x * x;
Console.WriteLine(square(5));
Console.WriteLine(square(8));
"#);
    assert_eq!(out, vec!["25", "64"]);
}

#[test]
fn func_two_args() {
    let out = run_csharp(r#"
Func<int, int, int> add = (a, b) => a + b;
Console.WriteLine(add(3, 4));
"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn action_no_params() {
    let out = run_csharp(r#"
Action sayHi = () => Console.WriteLine("hi");
sayHi();
"#);
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn lambda_closure_counter() {
    let out = run_csharp(r#"
int count = 0;
Action inc = () => { count++; };
inc();
inc();
inc();
Console.WriteLine(count);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn lambda_block_body() {
    let out = run_csharp(r#"
Func<int, string> classify = x => {
    if (x > 0) return "positive";
    if (x < 0) return "negative";
    return "zero";
};
Console.WriteLine(classify(5));
Console.WriteLine(classify(-3));
Console.WriteLine(classify(0));
"#);
    assert_eq!(out, vec!["positive", "negative", "zero"]);
}

#[test]
fn event_basic() {
    let out = run_csharp(r#"
class Button {
    public event Action Click;
    public void Press() {
        if (Click != null) Click();
    }
}
var btn = new Button();
btn.Click += () => Console.WriteLine("clicked!");
btn.Press();
btn.Press();
"#);
    assert_eq!(out, vec!["clicked!", "clicked!"]);
}

#[test]
fn delegate_declaration() {
    let out = run_csharp(r#"
delegate int MathOp(int a, int b);
class Program {
    public static int Add(int a, int b) { return a + b; }
    public static int Mul(int a, int b) { return a * b; }
}
MathOp op = (a, b) => a + b;
Console.WriteLine(op(3, 4));
op = (a, b) => a * b;
Console.WriteLine(op(3, 4));
"#);
    assert_eq!(out, vec!["7", "12"]);
}
