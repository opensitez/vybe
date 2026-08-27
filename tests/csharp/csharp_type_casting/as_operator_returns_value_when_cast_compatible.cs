// vybe-test: csharp/csharp_type_casting/as_operator_returns_value_when_cast_compatible
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

using static __Harness;

object o = "hello";
string s = o as string;
__P((s).ToString());
__Check("hello");

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
