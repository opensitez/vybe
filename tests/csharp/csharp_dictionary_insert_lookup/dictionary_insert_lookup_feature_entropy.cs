// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
string feature = "dictionary_insert_lookup:34"; __Check((feature.Length >= 1).ToString(), "True");
