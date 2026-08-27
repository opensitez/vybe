// vybe-test: csharp/csharp_generics/generic_method
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

using static __Harness;

__P((Utils.Identity<int>(42)).ToString());
__P((Utils.Identity<string>("hello")).ToString());
__Check("42\nhello");

class Utils {
    public static T Identity<T>(T value) { return value; }
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
