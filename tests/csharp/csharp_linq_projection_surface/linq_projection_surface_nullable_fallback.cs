// vybe-test: csharp/csharp_linq_projection_surface/linq_projection_surface_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_linq_projection_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_projection_surface
int? maybe = null; int fallback = maybe ?? 118; __Check((fallback == 118).ToString(), "True");
