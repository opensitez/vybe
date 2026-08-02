// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
string feature = "field_scope_surface"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
