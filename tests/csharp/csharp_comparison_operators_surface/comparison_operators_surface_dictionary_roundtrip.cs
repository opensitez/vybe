// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[13] = 14; __Check((map.ContainsKey(13) && map[13] == 14).ToString(), "True");
