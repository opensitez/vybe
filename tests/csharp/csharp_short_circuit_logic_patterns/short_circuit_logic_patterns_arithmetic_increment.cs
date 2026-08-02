// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
int seed = 14; __Check((seed + 1 > seed).ToString(), "True");
