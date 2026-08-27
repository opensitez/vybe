// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_last_batch_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

__P((new[]{1,2,3,4,5}.Chunk(2).Last().Sum()).ToString());
__Check("5");

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
