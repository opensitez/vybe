// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
string feature = "field_scope_surface"; __Check((feature[0] == feature[0]).ToString(), "True");
