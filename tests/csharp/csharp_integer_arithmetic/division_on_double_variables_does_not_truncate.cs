// vybe-test: csharp/csharp_integer_arithmetic/division_on_double_variables_does_not_truncate
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

using static __Harness;

double a = 9, b = 4;
__P((a / b).ToString());
__Check("2.25");

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
