// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
int? maybe = 67; __Check((maybe.HasValue && maybe.Value == 67).ToString(), "True");
