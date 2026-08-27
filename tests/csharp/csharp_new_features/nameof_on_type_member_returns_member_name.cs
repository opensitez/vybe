// vybe-test: csharp/csharp_new_features/nameof_on_type_member_returns_member_name
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

using static __Harness;

__P((nameof(Widget.Count)).ToString());
__Check("Count");

class Widget { public int Count; }

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
