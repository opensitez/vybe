// vybe-test: csharp/csharp_memory_patterns/array_pool_rent_returns_array_of_at_least_requested_size
// origin: languages/csharp/tests/csharp/test_csharp_memory_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pool=System.Buffers.ArrayPool<int>.Shared;
var arr=pool.Rent(10);
__Check((arr.Length>=10).ToString(), "True");
pool.Return(arr);
