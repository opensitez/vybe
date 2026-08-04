// vybe-test: csharp/csharp_fluent_builder_pattern/interface_typed_variable_can_invoke_fluent_concrete_methods
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

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
__P((writer.With("x").With("y").Join()).ToString());
__Check("xy");
