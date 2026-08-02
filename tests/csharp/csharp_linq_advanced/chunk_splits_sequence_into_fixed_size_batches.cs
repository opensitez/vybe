// vybe-test: csharp/csharp_linq_advanced/chunk_splits_sequence_into_fixed_size_batches
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var batches=new[]{1,2,3,4,5}.Chunk(2).ToList();
__Check((batches.Count).ToString(), "3");
__Check((batches[0].Length).ToString(), "2");
