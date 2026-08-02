// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// short_circuit_logic_patterns
int seed = 14; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
