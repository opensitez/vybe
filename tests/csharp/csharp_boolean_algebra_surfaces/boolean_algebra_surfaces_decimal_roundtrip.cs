// vybe-test: csharp/csharp_boolean_algebra_surfaces/boolean_algebra_surfaces_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_boolean_algebra_surfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boolean_algebra_surfaces
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
