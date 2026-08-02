// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
int seed = 37; __Check((seed - seed == 0).ToString(), "True");
