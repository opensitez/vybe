// vybe-test: csharp/csharp_floating_point_literals_surface/floating_point_literals_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_floating_point_literals_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// floating_point_literals_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[16] = 17; __Check((map.ContainsKey(16) && map[16] == 17).ToString(), "True");
