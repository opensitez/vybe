// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
