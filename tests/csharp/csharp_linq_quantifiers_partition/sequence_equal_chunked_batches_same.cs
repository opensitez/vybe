// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_chunked_batches_same
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

var a=new[]{1,2,3,4}
;
var b=new[]{1,2,3,4}
;
__P((a.Chunk(2).SelectMany(x=>x).SequenceEqual(b)).ToString());
__Check("True");

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
