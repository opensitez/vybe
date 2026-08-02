// vybe-test: csharp/csharp_bcl_collections/dictionary_with_string_comparer_uses_case_insensitive_key_lookup
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var map = new System.Collections.Generic.Dictionary<string, int>(
    System.StringComparer.OrdinalIgnoreCase);
map["Key"] = 7;
__Check((map.ContainsKey("key")).ToString(), "True");
