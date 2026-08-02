// vybe-test: csharp/csharp_collections_initialise/collection_expression_empty_array_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] empty=[];
__Check((empty.Length).ToString(), "0");
