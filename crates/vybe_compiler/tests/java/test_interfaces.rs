use crate::helpers::run_in_main;

#[test]
fn class_implements_interface_method() {
    let types = r#"
        interface Greeter { String greet(); }
        static class EnglishGreeter implements Greeter {
            public String greet() { return "hello"; }
        }
    "#;
    let out = run_in_main(
        "Greeter g = new EnglishGreeter(); System.out.println(g.greet());",
        types,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn interface_reference_dispatches_to_implementation() {
    let types = r#"
        interface Calc { int doubleIt(int n); }
        static class Doubler implements Calc {
            public int doubleIt(int n) { return n * 2; }
        }
    "#;
    let out = run_in_main(
        "Calc c = new Doubler(); System.out.println(c.doubleIt(6));",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn class_implements_multiple_interfaces() {
    let types = r#"
        interface A { default String fromA() { return "A"; } }
        interface B { default String fromB() { return "B"; } }
        static class Both implements A, B {}
    "#;
    let out = run_in_main(
        "Both b = new Both(); System.out.println(b.fromA() + b.fromB());",
        types,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn interface_default_method_used_when_not_overridden() {
    let types = r#"
        interface Logger { default void log(String msg) { System.out.println(msg); } }
        static class ConsoleLogger implements Logger {}
    "#;
    let out = run_in_main(
        "Logger l = new ConsoleLogger(); l.log(\"ok\");",
        types,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn interface_static_method_called_by_qualified_name() {
    let types = r#"
        interface MathUtil { static int triple(int n) { return n * 3; } }
    "#;
    let out = run_in_main(
        "System.out.println(MathUtil.triple(4));",
        types,
    );
    assert_eq!(out, vec!["12"]);
}
