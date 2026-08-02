// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
var map = new System.Collections.Generic.Dictionary<int, int>(); map[34] = 35; __Check((map.ContainsKey(34) && map[34] == 35).ToString(), "True");
