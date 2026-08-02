// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
int seed = 63; int right = seed + 1; __Check((seed < right).ToString(), "True");
