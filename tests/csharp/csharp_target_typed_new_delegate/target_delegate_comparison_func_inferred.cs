// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_comparison_func_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Func<int, int, bool> less = (a, b) => a < b;
__P((less(2, 5)).ToString());
__P((less(9, 1)).ToString());
__Check("True\nFalse");

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
