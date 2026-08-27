// vybe-test: csharp/csharp_collection_expressions/collection_expression_long_array_sum_loop
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

long[] nums = [10000000000L, 20000000000L];
long total = 0;
foreach (var n in nums) total += n;
__P((total).ToString());
__Check("30000000000");

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
