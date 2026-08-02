// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_chunked_batches_same
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new[]{1,2,3,4};
var b=new[]{1,2,3,4};
__Check((a.Chunk(2).SelectMany(x=>x).SequenceEqual(b)).ToString(), "True");
