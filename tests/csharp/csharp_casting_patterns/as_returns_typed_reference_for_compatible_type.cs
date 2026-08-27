// vybe-test: csharp/csharp_casting_patterns/as_returns_typed_reference_for_compatible_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

using static __Harness;

object o="world";
string s=o as string;
__P((s!=null).ToString());
__P((s).ToString());
__Check("True\nworld");

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
