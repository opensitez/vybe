// vybe-test: csharp/csharp_null_propagation/null_conditional_indexer_uses_fallback_for_null_array
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] values = null; __Check((values?[0] ?? -1).ToString(), "-1");
