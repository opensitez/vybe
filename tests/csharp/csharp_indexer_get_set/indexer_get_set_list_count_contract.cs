// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
var values = new System.Collections.Generic.List<int> { 66, 67, 66 }; __Check((values.Count == 3).ToString(), "True");
