// vybe-test: csharp/csharp_memory_patterns/array_pool_rent_returns_array_of_at_least_requested_size
// origin: languages/csharp/tests/csharp/test_csharp_memory_patterns.rs

using static __Harness;

var pool=System.Buffers.ArrayPool<int>.Shared;
var arr=pool.Rent(10);
__P((arr.Length>=10).ToString());
pool.Return(arr);
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
