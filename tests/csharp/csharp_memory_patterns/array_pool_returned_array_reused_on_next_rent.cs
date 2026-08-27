// vybe-test: csharp/csharp_memory_patterns/array_pool_returned_array_reused_on_next_rent
// origin: languages/csharp/tests/csharp/test_csharp_memory_patterns.rs

using static __Harness;

var pool=System.Buffers.ArrayPool<byte>.Shared;
var a=pool.Rent(8);
pool.Return(a,clearArray:true);
var b=pool.Rent(8);
__P((b.Length>=8).ToString());
pool.Return(b);
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
