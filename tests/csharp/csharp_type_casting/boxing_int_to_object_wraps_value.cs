// vybe-test: csharp/csharp_type_casting/boxing_int_to_object_wraps_value
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

using static __Harness;

object boxed = 42;
__P((boxed).ToString());
__Check("42");

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
