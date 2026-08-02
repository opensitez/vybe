// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// break_continue_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[49] = 50; __Check((map.ContainsKey(49) && map[49] == 50).ToString(), "True");
