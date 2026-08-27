// vybe-test: csharp/interfaces_generics/extension_method_on_list
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

var nums = new List<int> { 1, 2, 3, 4, 5 }
;
__P((nums.Join(", ")).ToString());
__Check("1, 2, 3, 4, 5");

static class ListExtensions {
    public static string Join<T>(this List<T> list, string sep) {
        return string.Join(sep, list);
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
