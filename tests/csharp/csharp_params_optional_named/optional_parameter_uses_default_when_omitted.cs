// vybe-test: csharp/csharp_params_optional_named/optional_parameter_uses_default_when_omitted
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

using static __Harness;

string Greet(string name, string prefix="Hello") => prefix+" "+name;
__P((Greet("World")).ToString());
__Check("Hello World");

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
