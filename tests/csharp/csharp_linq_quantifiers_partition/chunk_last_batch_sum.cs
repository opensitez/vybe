// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_last_batch_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2,3,4,5}.Chunk(2).Last().Sum()).ToString(), "5");
