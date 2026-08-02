// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
int seed = 37; int right = seed + 1; __Check((seed < right).ToString(), "True");
