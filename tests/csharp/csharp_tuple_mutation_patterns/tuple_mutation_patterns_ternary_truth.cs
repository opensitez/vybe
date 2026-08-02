// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
int seed = 37; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
