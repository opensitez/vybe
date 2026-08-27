// vybe-test: csharp/interfaces_generics/extension_method_basic
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

__P(("hello".Reverse()).ToString());
__Check("olleh");

static class StringExtensions {
    public static string Reverse(this string s) {
        char[] chars = s.ToCharArray();
        Array.Reverse(chars);
        return new string(chars);
    }
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
