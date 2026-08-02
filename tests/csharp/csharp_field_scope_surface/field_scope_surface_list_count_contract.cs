// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
var values = new System.Collections.Generic.List<int> { 63, 64, 63 }; __Check((values.Count == 3).ToString(), "True");
