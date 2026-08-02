// vybe-test: csharp/csharp_linq_quantifiers_partition/all_predicate_on_chunk_first_batch
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new[]{2,4,6,8}.Chunk(2).First().All(x=>x%2==0)).ToString(), "True");
