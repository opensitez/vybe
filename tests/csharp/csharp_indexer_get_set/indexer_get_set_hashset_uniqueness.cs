// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
var set = new System.Collections.Generic.HashSet<int>(); set.Add(66); set.Add(66); __Check((set.Count == 1).ToString(), "True");
