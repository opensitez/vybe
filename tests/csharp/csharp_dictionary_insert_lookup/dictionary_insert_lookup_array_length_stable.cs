// vybe-test: csharp/csharp_dictionary_insert_lookup/dictionary_insert_lookup_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_insert_lookup.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// dictionary_insert_lookup
int seed = 34; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
