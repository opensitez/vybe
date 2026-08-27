// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_then_all_batches_full_except_last
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

var batches=new[]{1,2,3,4,5,6,7}
.Chunk(3);
__P((batches.Take(2).All(b=>b.Length==3)?1:0).ToString());
__P((batches.Last().Length).ToString());
__Check("1\n1");

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
