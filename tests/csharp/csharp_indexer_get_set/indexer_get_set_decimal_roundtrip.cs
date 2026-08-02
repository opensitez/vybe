// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
