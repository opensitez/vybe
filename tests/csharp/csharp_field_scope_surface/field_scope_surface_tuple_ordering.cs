// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
var tuple = (left: 63, right: 64); __Check((tuple.left < tuple.right).ToString(), "True");
