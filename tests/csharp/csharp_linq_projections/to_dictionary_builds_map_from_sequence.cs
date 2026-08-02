// vybe-test: csharp/csharp_linq_projections/to_dictionary_builds_map_from_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dict = new[]{"a","bb","ccc"}.ToDictionary(s => s, s => s.Length);
__Check((dict["bb"]).ToString(), "2");
