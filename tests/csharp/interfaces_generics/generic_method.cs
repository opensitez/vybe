// vybe-test: csharp/interfaces_generics/generic_method
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

__P((Utils.Max(3, 7)).ToString());
__P((Utils.Max("apple", "banana")).ToString());
__Check("7\nbanana");

class Utils {
    public static T Max<T>(T a, T b) where T : IComparable<T> {
        return a.CompareTo(b) > 0 ? a : b;
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
