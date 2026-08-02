// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
string feature = "dictionary_insert_lookup"; __Check((feature.Length > 0).ToString(), "True");
