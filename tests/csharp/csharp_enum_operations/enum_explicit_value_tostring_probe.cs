// vybe-test: csharp/csharp_enum_operations/enum_explicit_value_tostring_probe
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

using static __Harness;

__P((Num.X.ToString()).ToString());
__Check("X");

enum Num{X=7}

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
