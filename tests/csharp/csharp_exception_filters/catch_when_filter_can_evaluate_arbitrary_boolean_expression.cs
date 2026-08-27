// vybe-test: csharp/csharp_exception_filters/catch_when_filter_can_evaluate_arbitrary_boolean_expression
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

using static __Harness;

int threshold = 10;
try {
    throw new System.InvalidOperationException("value=15");
}
catch (System.InvalidOperationException e) when (threshold < 20) {
    __P(("caught with threshold").ToString());
}
__Check("caught with threshold");

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
