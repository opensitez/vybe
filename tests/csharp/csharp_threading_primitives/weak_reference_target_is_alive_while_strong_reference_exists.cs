// vybe-test: csharp/csharp_threading_primitives/weak_reference_target_is_alive_while_strong_reference_exists
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

using static __Harness;

var strong = new object();
var weak = new System.WeakReference(strong);
__P((weak.IsAlive).ToString());
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
