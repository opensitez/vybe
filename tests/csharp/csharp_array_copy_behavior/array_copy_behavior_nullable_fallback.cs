// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

using static __Harness;

// array_copy_behavior
int? maybe = null;
int fallback = maybe ?? 26;
__P((fallback == 26).ToString());
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
