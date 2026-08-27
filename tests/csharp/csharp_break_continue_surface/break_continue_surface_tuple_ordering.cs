// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

using static __Harness;

// break_continue_surface
var tuple = (left: 49, right: 50);
__P((tuple.left < tuple.right).ToString());
__Check("True");

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
