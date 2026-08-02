// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
int? maybe = 11; __Check((maybe.HasValue && maybe.Value == 11).ToString(), "True");
