// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
int seed = 63; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
