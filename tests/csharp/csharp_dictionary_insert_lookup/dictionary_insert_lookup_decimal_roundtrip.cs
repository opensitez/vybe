// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
