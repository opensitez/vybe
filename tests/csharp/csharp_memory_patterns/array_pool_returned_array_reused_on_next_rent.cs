// vybe-test: csharp/csharp_memory_patterns/array_pool_returned_array_reused_on_next_rent
// origin: languages/csharp/tests/csharp/test_csharp_memory_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pool=System.Buffers.ArrayPool<byte>.Shared;
var a=pool.Rent(8);
pool.Return(a,clearArray:true);
var b=pool.Rent(8);
__Check((b.Length>=8).ToString(), "True");
pool.Return(b);
