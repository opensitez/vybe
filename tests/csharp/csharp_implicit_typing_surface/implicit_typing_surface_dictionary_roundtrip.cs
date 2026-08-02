// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// implicit_typing_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[59] = 60; __Check((map.ContainsKey(59) && map[59] == 60).ToString(), "True");
