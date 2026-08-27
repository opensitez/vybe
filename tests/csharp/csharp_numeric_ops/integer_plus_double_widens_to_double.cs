// vybe-test: csharp/csharp_numeric_ops/integer_plus_double_widens_to_double
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

using static __Harness;

int i=3;
double d=1.5;
__P((i+d).ToString());
__Check("4.5");

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
