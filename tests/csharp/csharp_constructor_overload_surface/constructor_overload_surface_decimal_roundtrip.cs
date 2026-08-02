// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
