// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
