// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
double seed = 37; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
