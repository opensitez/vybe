// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
int? maybe = 34; __Check((maybe.HasValue && maybe.Value == 34).ToString(), "True");
