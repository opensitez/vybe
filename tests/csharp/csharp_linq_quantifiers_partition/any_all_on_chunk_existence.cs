// vybe-test: csharp/csharp_linq_quantifiers_partition/any_all_on_chunk_existence
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var batches=new[]{1,2,3,4}.Chunk(2);
__Check((batches.Any()?1:0).ToString(), "1");
__Check((batches.All(b=>b.Length>0)?1:0).ToString(), "1");
