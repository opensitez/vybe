// vybe-test: csharp/csharp_linq_numeric/count_with_predicate_counts_matching
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

using static __Harness;

__P((new[]{1,2,3,4,5,6}.Count(n=>n%2==0)).ToString());
__Check("3");

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
