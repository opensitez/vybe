// vybe-test: csharp/csharp_linq_advanced/chunk_splits_sequence_into_fixed_size_batches
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

using static __Harness;

var batches=new[]{1,2,3,4,5}
.Chunk(2).ToList();
__P((batches.Count).ToString());
__P((batches[0].Length).ToString());
__Check("3\n2");

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
