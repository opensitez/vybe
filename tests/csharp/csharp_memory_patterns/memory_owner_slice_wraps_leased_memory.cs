// vybe-test: csharp/csharp_memory_patterns/memory_owner_slice_wraps_leased_memory
// origin: languages/csharp/tests/csharp/test_csharp_memory_patterns.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using var owner=System.Buffers.MemoryPool<int>.Shared.Rent(5);
var span=owner.Memory.Span;
for(int i=0;i<5;i++) span[i]=i+1;
__P((span[4]).ToString());
__Check("5");
