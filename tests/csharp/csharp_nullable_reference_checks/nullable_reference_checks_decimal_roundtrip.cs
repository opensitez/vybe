// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
