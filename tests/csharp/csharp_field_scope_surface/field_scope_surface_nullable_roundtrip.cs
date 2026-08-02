// vybe-test: csharp/csharp_field_scope_surface/field_scope_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_field_scope_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// field_scope_surface
int? maybe = 63; __Check((maybe.HasValue && maybe.Value == 63).ToString(), "True");
