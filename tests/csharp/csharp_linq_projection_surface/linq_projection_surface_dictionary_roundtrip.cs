// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[118] = 119; __Check((map.ContainsKey(118) && map[118] == 119).ToString(), "True");
