// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
