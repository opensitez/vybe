// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
double seed = 27; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
