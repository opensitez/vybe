// vybe-test: csharp/csharp_using_static/using_static_enumerable_allows_range
// origin: languages/csharp/tests/csharp/test_csharp_using_static.rs

using static __Harness;
using static System.Linq.Enumerable;

__P((string.Join(",",Range(1,4))).ToString());
__Check("1,2,3,4");

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
