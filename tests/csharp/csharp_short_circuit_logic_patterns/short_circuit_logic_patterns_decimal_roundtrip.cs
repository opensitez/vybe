// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
