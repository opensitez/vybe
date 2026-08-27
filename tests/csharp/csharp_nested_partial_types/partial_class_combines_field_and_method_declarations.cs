// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_field_and_method_declarations
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

__P((new Config().Read()).ToString());
__Check("prod");

partial class Config {
    string env = "prod";
}

partial class Config {
    public string Read() { return env; }
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
