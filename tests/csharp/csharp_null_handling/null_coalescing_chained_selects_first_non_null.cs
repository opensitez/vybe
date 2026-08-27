// vybe-test: csharp/csharp_null_handling/null_coalescing_chained_selects_first_non_null
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

using static __Harness;

string a=null, b=null, c="found";
__P((a ?? b ?? c).ToString());
__Check("found");

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
