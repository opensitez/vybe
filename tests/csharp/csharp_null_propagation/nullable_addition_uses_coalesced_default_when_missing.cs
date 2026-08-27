// vybe-test: csharp/csharp_null_propagation/nullable_addition_uses_coalesced_default_when_missing
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;

int? left = null;
int? right = 5;
__P(((left ?? 0) + (right ?? 0)).ToString());
__Check("5");

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
