// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
var values = new System.Collections.Generic.List<int> { 34, 35, 34 }; __Check((values.Count == 3).ToString(), "True");
