// vybe-test: csharp/csharp_threading_primitives/lazy_is_value_created_flips_after_materialization
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

using static __Harness;

var lazy = new System.Lazy<int>(() => 3);
__P((lazy.IsValueCreated).ToString());
__P((lazy.Value).ToString());
__P((lazy.IsValueCreated).ToString());
__Check("False\n3\nTrue");

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
