// vybe-test: csharp/csharp_string_interpolation/ternary_operator_inside_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

using static __Harness;

int n=7;
__P(($"{(n%2==0?"even":"odd")}").ToString());
__Check("odd");

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
