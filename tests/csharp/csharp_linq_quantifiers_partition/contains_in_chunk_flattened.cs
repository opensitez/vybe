// vybe-test: csharp/csharp_linq_quantifiers_partition/contains_in_chunk_flattened
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{1,2,3,4,5}.Chunk(2).SelectMany(x=>x);
__Check((flat.Contains(5)?1:0).ToString(), "1");
