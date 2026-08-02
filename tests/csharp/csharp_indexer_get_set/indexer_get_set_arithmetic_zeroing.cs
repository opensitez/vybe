// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
int seed = 66; __Check((seed - seed == 0).ToString(), "True");
