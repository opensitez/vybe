// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

using static __Harness;

// static_constructor_guard
int seed = 69;
int right = seed + 1;
__P((seed < right).ToString());
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
