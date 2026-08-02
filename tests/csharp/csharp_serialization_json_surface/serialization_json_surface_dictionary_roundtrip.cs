// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// serialization_json_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[91] = 92; __Check((map.ContainsKey(91) && map[91] == 92).ToString(), "True");
