// vybe-test: csharp/csharp_linq_aggregates/aggregate_folds_sequence_with_seed_and_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

using static __Harness;

__P((new[]{1,2,3,4}.Aggregate(0, (acc, x) => acc + x)).ToString());
__Check("10");

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
