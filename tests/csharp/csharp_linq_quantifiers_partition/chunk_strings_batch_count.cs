// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_strings_batch_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{"a","b","c","d","e"}.Chunk(2).Count()).ToString(), "3");
