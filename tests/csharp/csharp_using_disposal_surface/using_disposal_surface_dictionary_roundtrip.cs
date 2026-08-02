// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[52] = 53; __Check((map.ContainsKey(52) && map[52] == 53).ToString(), "True");
