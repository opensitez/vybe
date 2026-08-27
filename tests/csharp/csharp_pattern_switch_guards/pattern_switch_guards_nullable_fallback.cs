// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

using static __Harness;

// pattern_switch_guards
int? maybe = null;
int fallback = maybe ?? 42;
__P((fallback == 42).ToString());
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
