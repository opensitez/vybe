// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
int seed = 27; __Check((seed + 1 > seed).ToString(), "True");
