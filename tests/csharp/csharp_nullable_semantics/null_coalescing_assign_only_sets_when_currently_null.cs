// vybe-test: csharp/csharp_nullable_semantics/null_coalescing_assign_only_sets_when_currently_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

using static __Harness;

int? n = null;
n ??= 5;
__P((n).ToString());
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
