// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
var map = new System.Collections.Generic.Dictionary<int, int>(); map[65] = 66; __Check((map.ContainsKey(65) && map[65] == 66).ToString(), "True");
