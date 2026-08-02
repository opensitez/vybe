// vybe-test: csharp/csharp_explicit_typing_surface/explicit_typing_surface_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_explicit_typing_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// explicit_typing_surface
int? maybe = null; int fallback = maybe ?? 60; __Check((fallback == 60).ToString(), "True");
