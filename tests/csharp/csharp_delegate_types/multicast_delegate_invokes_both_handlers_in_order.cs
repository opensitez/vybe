// vybe-test: csharp/csharp_delegate_types/multicast_delegate_invokes_both_handlers_in_order
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

using static __Harness;

System.Action log = () => __P(("a").ToString());
log += () => __P(("b").ToString());
log();
__Check("a\nb");

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
