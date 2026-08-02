// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_then_all_batches_full_except_last
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var batches=new[]{1,2,3,4,5,6,7}.Chunk(3);
__Check((batches.Take(2).All(b=>b.Length==3)?1:0).ToString(), "1");
__Check((batches.Last().Length).ToString(), "1");
