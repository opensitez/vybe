// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
int? maybe = 14; __Check((maybe.HasValue && maybe.Value == 14).ToString(), "True");
