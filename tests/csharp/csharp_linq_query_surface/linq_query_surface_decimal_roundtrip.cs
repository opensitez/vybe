// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_query_surface
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
