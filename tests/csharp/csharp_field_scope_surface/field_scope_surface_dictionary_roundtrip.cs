// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[63] = 64; __Check((map.ContainsKey(63) && map[63] == 64).ToString(), "True");
