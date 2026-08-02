// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
