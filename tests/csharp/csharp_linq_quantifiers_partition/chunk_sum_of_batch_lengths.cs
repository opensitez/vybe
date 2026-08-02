// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_sum_of_batch_lengths
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var batches=new[]{1,2,3,4,5,6}.Chunk(2);
__Check((batches.Sum(b=>b.Length)).ToString(), "6");
