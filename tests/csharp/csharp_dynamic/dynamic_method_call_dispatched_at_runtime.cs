// vybe-test: csharp/csharp_dynamic/dynamic_method_call_dispatched_at_runtime
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

using static __Harness;

object o="hello";
dynamic d=o;
__P((d.ToUpper()).ToString());
__Check("HELLO");

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
