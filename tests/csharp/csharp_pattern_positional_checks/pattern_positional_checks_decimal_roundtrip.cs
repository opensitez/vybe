// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
