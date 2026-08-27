// vybe-test: csharp/csharp_exceptions_hierarchy/aggregate_exception_wraps_multiple_inner_exceptions
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

using static __Harness;

var ae=new System.AggregateException(
    new System.Exception("one"),
    new System.Exception("two"));
__P((ae.InnerExceptions.Count).ToString());
__Check("2");

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
