// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[67] = 68; __Check((map.ContainsKey(67) && map[67] == 68).ToString(), "True");
