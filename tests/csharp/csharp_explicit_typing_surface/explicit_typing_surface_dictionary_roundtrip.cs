// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// explicit_typing_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[60] = 61; __Check((map.ContainsKey(60) && map[60] == 61).ToString(), "True");
