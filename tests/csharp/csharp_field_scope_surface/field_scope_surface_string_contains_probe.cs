// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
string feature = "field_scope_surface"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
