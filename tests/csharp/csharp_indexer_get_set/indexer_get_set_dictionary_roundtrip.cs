// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// indexer_get_set
var map = new System.Collections.Generic.Dictionary<int, int>(); map[66] = 67; __Check((map.ContainsKey(66) && map[66] == 67).ToString(), "True");
