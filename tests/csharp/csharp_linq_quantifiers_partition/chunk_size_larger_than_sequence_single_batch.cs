// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_size_larger_than_sequence_single_batch
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{1,2}.Chunk(5).Count()).ToString(), "1");
