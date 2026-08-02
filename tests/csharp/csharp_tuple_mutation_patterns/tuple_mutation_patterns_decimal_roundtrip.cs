// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
