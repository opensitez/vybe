// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
int seed = 14; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
