// vybe-test: csharp/csharp_nameof_expressions/nameof_interface_type_returns_type_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

using static __Harness;

string n = nameof(SampleClass);
__P(n);
__Check("SampleClass");

class SampleClass { }
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
