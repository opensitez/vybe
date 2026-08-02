// vybe-test: csharp/csharp_collections_initialise/collection_expression_creates_array_directly
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr=[10,20,30];
__Check((arr.Length).ToString(), "3");
