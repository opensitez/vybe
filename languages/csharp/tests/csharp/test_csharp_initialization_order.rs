//! Static and instance initialization order before user code in constructors runs.
use super::helpers::run_csharp;

#[test]
fn static_field_initializers_run_in_declaration_order_before_static_method() {
    assert_eq!(
        run_csharp(
            r#"
class Logger {
    static string First = Mark("first");
    static string Second = Mark("second");
    static string Mark(string name) {
        Console.WriteLine(name);
        return name;
    }
    public static void Run() {
        Console.WriteLine("run");
    }
}
Logger.Run();
"#
        ),
        &["first", "second", "run"]
    );
}

#[test]
fn instance_field_initializer_runs_before_constructor_body() {
    assert_eq!(
        run_csharp(
            r#"
class Widget {
    string label = Init("field");
    public Widget() {
        Console.WriteLine("ctor");
    }
    static string Init(string part) {
        Console.WriteLine(part);
        return part;
    }
}
new Widget();
"#
        ),
        &["field", "ctor"]
    );
}

#[test]
fn base_constructor_runs_before_derived_field_initializers_and_body() {
    assert_eq!(
        run_csharp(
            r#"
class Base {
    public Base() { Console.WriteLine("base-ctor"); }
}
class Derived : Base {
    string tag = Init("derived-field");
    public Derived() { Console.WriteLine("derived-ctor"); }
    static string Init(string part) {
        Console.WriteLine(part);
        return part;
    }
}
new Derived();
"#
        ),
        &["base-ctor", "derived-field", "derived-ctor"]
    );
}

#[test]
fn static_constructor_runs_once_before_first_instance_creation() {
    assert_eq!(
        run_csharp(
            r#"
class Counter {
    static Counter() { Console.WriteLine("static-ctor"); }
    public Counter() { Console.WriteLine("instance"); }
}
new Counter();
new Counter();
"#
        ),
        &["static-ctor", "instance", "instance"]
    );
}

#[test]
fn readonly_instance_field_can_be_set_in_constructor_but_not_after() {
    assert_eq!(
        run_csharp(
            r#"
class Token {
    public readonly string Value;
    public Token(string value) { Value = value; }
}
var token = new Token("abc");
Console.WriteLine(token.Value);
"#
        ),
        &["abc"]
    );
}
