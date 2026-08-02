// vybe-test: csharp/csharp_memory_patterns/memory_owner_slice_wraps_leased_memory
// origin: languages/csharp/tests/csharp/test_csharp_memory_patterns.rs

using var owner=System.Buffers.MemoryPool<int>.Shared.Rent(5);
var span=owner.Memory.Span;
for(int i=0;i<5;i++) span[i]=i+1;
Console.WriteLine(span[4]);
