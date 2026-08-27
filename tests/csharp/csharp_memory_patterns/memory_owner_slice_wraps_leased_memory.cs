// vybe-test: csharp/csharp_memory_patterns/memory_owner_slice_wraps_leased_memory
// origin: languages/csharp/tests/csharp/test_csharp_memory_patterns.rs

using static __Harness;
using var owner=System.Buffers.MemoryPool<int>.Shared.Rent(5);

var span=owner.Memory.Span;
for(int i=0;i<5;i++) span[i]=i+1;
__P((span[4]).ToString());
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
