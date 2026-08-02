// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
string feature = "array_reverse_patterns:27"; __Check((feature.Length >= 1).ToString(), "True");
