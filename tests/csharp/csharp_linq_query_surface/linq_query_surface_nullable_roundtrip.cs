// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
int? maybe = 117; __Check((maybe.HasValue && maybe.Value == 117).ToString(), "True");
