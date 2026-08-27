// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_nullable_int_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Func<int?, int> orZero = n => n ?? 0;
__P((orZero(null)).ToString());
__P((orZero(7)).ToString());
__Check("0\n7");

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
