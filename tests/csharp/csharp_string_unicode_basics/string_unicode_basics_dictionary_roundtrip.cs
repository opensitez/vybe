// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
var map = new System.Collections.Generic.Dictionary<int, int>(); map[19] = 20; __Check((map.ContainsKey(19) && map[19] == 20).ToString(), "True");
