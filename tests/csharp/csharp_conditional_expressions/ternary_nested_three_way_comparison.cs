// vybe-test: csharp/csharp_conditional_expressions/ternary_nested_three_way_comparison
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

using static __Harness;

int n=0;
__P((n>0?"pos":n<0?"neg":"zero").ToString());
__Check("zero");

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
