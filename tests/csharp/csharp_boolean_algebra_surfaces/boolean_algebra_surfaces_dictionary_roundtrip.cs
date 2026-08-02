// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
var map = new System.Collections.Generic.Dictionary<int, int>(); map[11] = 12; __Check((map.ContainsKey(11) && map[11] == 12).ToString(), "True");
