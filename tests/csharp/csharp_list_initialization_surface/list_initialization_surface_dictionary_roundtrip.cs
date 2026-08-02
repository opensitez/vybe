// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// list_initialization_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[30] = 31; __Check((map.ContainsKey(30) && map[30] == 31).ToString(), "True");
