// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
int seed = 28; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
