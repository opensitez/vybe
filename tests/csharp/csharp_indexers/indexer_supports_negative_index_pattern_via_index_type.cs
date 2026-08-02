// vybe-test: csharp/csharp_indexers/indexer_supports_negative_index_pattern_via_index_type
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,2,3,4,5};
__Check((arr[^1]).ToString(), "5"); __Check((arr[^2]).ToString(), "4");
