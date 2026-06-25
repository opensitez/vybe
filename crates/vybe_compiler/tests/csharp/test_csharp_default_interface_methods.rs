//! Default interface methods provide callable implementation on the interface type.
use super::helpers::run_csharp;

#[test]
fn default_interface_method_invoked_on_concrete_type_without_override() {
    assert_eq!(
        run_csharp(
            r#"
interface IReporter {
    void Write(string text);
    void Banner(string title) { Write("==" + title + "=="); }
}
class ConsoleReporter : IReporter {
    public void Write(string text) { Console.WriteLine(text); }
}
var reporter = new ConsoleReporter();
reporter.Banner("start");
"#
        ),
        &["==start=="]
    );
}

#[test]
fn default_interface_method_visible_through_interface_typed_reference() {
    assert_eq!(
        run_csharp(
            r#"
interface ICounter {
    int Value { get; }
    int Next() { return Value + 1; }
}
class Counter : ICounter {
    public int Value { get; set; }
}
ICounter counter = new Counter { Value = 4 };
Console.WriteLine(counter.Next());
"#
        ),
        &["5"]
    );
}

#[test]
fn class_override_replaces_default_interface_method_implementation() {
    assert_eq!(
        run_csharp(
            r#"
interface IFormat {
    string Format(int n);
    string Label(int n) { return "d:" + Format(n); }
}
class Custom : IFormat {
    public string Format(int n) { return n.ToString(); }
    public string Label(int n) { return "x:" + Format(n); }
}
IFormat fmt = new Custom();
Console.WriteLine(fmt.Label(3));
"#
        ),
        &["x:3"]
    );
}
