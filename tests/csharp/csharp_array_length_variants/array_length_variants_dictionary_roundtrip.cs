// vybe-test: csharp/csharp_array_length_variants/array_length_variants_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
var map = new System.Collections.Generic.Dictionary<int, int>(); map[25] = 26; __Check((map.ContainsKey(25) && map[25] == 26).ToString(), "True");
