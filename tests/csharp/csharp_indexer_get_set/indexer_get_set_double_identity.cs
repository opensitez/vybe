// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
double seed = 66; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
