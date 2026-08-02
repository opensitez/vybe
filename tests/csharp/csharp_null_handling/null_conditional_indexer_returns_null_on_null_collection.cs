// vybe-test: csharp/csharp_null_handling/null_conditional_indexer_returns_null_on_null_collection
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = null;
__Check((arr?[0] ?? -1).ToString(), "-1");
