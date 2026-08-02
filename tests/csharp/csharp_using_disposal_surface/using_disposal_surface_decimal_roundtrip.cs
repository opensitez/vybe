// vybe-test: csharp/csharp_using_disposal_surface/using_disposal_surface_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// using_disposal_surface
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
