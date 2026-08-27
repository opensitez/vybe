// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

using static __Harness;

// inheritance_virtual_dispatch
int seed = 71;
__P((seed - seed == 0).ToString());
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
