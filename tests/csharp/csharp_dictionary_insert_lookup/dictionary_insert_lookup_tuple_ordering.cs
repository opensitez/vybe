// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
var tuple = (left: 34, right: 35); __Check((tuple.left < tuple.right).ToString(), "True");
