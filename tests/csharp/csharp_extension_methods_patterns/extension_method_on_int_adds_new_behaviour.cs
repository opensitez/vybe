// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_int_adds_new_behaviour
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

using static __Harness;

__P((4.IsEven()).ToString());
__P((3.IsEven()).ToString());
__Check("True\nFalse");

static class IntExt { public static bool IsEven(this int n) => n%2==0; }

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
