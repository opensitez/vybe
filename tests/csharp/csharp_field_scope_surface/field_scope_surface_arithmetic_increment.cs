// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
int seed = 63; __Check((seed + 1 > seed).ToString(), "True");
