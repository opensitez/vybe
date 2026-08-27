// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

using static __Harness;

// multidim_array_traversal
int? maybe = null;
int fallback = maybe ?? 29;
__P((fallback == 29).ToString());
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
