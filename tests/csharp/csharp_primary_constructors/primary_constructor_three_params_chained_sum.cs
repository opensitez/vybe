// vybe-test: csharp/csharp_primary_constructors/primary_constructor_three_params_chained_sum
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Triple(1, 2, 3).Total).ToString());
__Check("6");

class Triple(int a, int b, int c) { public int Total => a + b + c; }

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
