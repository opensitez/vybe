// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
int? maybe = null; int fallback = maybe ?? 11; __Check((fallback == 11).ToString(), "True");
