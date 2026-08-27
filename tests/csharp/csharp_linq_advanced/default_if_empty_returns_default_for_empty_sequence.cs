// vybe-test: csharp/csharp_linq_advanced/default_if_empty_returns_default_for_empty_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

using static __Harness;

var result=System.Array.Empty<int>().DefaultIfEmpty(99);
__P((result.First()).ToString());
__Check("99");

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
