//! Fluent APIs that return `this` chain mutations and preserve object identity.
use super::helpers::run_csharp;

#[test]
fn chained_instance_methods_return_same_object_identity() {
    assert_eq!(
        run_csharp(
            r#"
class Builder {
    static int nextId;
    int id = ++nextId;
    int total;
    public Builder Add(int value) { total += value; return this; }
    public int Id() { return id; }
    public int Build() { return total; }
}
var builder = new Builder();
var same = builder.Add(2).Add(3);
Console.WriteLine(same.Id() == builder.Id() ? "Y" : "N");
Console.WriteLine(builder.Build());
"#
        ),
        &["Y", "5"]
    );
}

#[test]
fn fluent_chain_order_matches_call_sequence() {
    assert_eq!(
        run_csharp(
            r#"
class Trace {
    string log = "";
    public Trace Step(string name) { log += name; return this; }
    public string Read() { return log; }
}
var trace = new Trace().Step("a").Step("b").Step("c");
Console.WriteLine(trace.Read());
"#
        ),
        &["abc"]
    );
}

#[test]
fn static_factory_method_can_start_chain_on_new_instance() {
    assert_eq!(
        run_csharp(
            r#"
class Counter {
    int value;
    public static Counter Start(int seed) {
        var counter = new Counter();
        counter.value = seed;
        return counter;
    }
    public Counter Bump() { value++; return this; }
    public int Read() { return value; }
}
Console.WriteLine(Counter.Start(10).Bump().Bump().Read());
"#
        ),
        &["12"]
    );
}

#[test]
fn interface_typed_variable_can_invoke_fluent_concrete_methods() {
    assert_eq!(
        run_csharp(
            r#"
interface IAppend {
    IAppend With(string part);
    string Join();
}
class Joiner : IAppend {
    string text = "";
    public IAppend With(string part) { text += part; return this; }
    public string Join() { return text; }
}
IAppend writer = new Joiner();
Console.WriteLine(writer.With("x").With("y").Join());
"#
        ),
        &["xy"]
    );
}
