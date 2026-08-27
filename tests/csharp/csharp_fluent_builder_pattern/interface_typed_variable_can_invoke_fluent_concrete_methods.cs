// vybe-test: csharp/csharp_fluent_builder_pattern/interface_typed_variable_can_invoke_fluent_concrete_methods
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

using static __Harness;

IAppend writer = new Joiner();
__P((writer.With("x").With("y").Join()).ToString());
__Check("xy");

interface IAppend {
    IAppend With(string part);
    string Join();
}

class Joiner : IAppend {
    string text = "";
    public IAppend With(string part) { text += part; return this; }
    public string Join() { return text; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
