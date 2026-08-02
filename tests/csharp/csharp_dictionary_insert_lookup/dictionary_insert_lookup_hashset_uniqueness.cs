// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
var set = new System.Collections.Generic.HashSet<int>(); set.Add(34); set.Add(34); __Check((set.Count == 1).ToString(), "True");
