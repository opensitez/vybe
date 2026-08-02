// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
int seed = 27; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
