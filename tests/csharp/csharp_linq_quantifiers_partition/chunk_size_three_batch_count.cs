// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_size_three_batch_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3,4,5,6,7}.Chunk(3).Count()).ToString(), "3");
