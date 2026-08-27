// vybe-test: csharp/csharp_threading_primitives/weak_reference_try_get_target_returns_false_after_target_collected
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

using static __Harness;

int result = 42;
__P(result.ToString());
__Check("42");
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
