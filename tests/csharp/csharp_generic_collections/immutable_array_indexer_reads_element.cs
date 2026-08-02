// vybe-test: csharp/csharp_generic_collections/immutable_array_indexer_reads_element
// origin: languages/csharp/tests/csharp/test_csharp_generic_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr = System.Collections.Immutable.ImmutableArray.Create(10,20,30);
__Check((arr[1]).ToString(), "20");
