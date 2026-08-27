// vybe-test: csharp/interfaces_generics/extension_method_on_int
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

__P((4.IsEven()).ToString());
__P((3.IsEven()).ToString());
__P((5.Square()).ToString());
__Check("True\nFalse\n25");

static class IntExtensions {
    public static bool IsEven(this int n) { return n % 2 == 0; }
    public static int Square(this int n) { return n * n; }
}

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
