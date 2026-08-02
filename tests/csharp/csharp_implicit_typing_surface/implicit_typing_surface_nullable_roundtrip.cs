// vybe-test: csharp/csharp_implicit_typing_surface/implicit_typing_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_implicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// implicit_typing_surface
int? maybe = 59; __Check((maybe.HasValue && maybe.Value == 59).ToString(), "True");
