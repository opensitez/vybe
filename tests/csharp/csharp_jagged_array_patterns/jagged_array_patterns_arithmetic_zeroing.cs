// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
int seed = 28; __Check((seed - seed == 0).ToString(), "True");
