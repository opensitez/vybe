// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
int? maybe = null; int fallback = maybe ?? 66; __Check((fallback == 66).ToString(), "True");
