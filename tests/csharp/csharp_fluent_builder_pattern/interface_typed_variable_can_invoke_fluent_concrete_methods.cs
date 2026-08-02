// vybe-test: csharp/csharp_fluent_builder_pattern/interface_typed_variable_can_invoke_fluent_concrete_methods
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((writer.With("x").With("y").Join()).ToString(), "xy");
