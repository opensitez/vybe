// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_returning_delegate_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Func<int, System.Func<int, int>> scale = factor => n => n * factor;
var triple = scale(3);
__P((triple(4)).ToString());
__Check("12");

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
