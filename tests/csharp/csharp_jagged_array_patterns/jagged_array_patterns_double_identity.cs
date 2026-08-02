// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
double seed = 28; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
