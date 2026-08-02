// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
var tuple = (left: 66, right: 67); __Check((tuple.left < tuple.right).ToString(), "True");
