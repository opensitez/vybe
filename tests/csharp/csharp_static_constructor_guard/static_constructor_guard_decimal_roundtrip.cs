// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
