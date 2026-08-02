// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
string feature = "tuple_mutation_patterns"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
