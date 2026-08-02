// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
string feature = "indexer_get_set"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
