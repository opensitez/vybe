// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
int? maybe = null; int fallback = maybe ?? 52; __Check((fallback == 52).ToString(), "True");
