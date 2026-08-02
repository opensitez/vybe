// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
var map = new System.Collections.Generic.Dictionary<int, int>(); map[20] = 21; __Check((map.ContainsKey(20) && map[20] == 21).ToString(), "True");
