// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
string feature = "tuple_mutation_patterns:37"; __Check((feature.Length >= 1).ToString(), "True");
