// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
int seed = 66; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
