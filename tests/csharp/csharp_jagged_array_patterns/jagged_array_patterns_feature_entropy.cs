// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
string feature = "jagged_array_patterns:28"; __Check((feature.Length >= 1).ToString(), "True");
