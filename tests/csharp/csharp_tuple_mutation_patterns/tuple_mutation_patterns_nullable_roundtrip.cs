// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
int? maybe = 37; __Check((maybe.HasValue && maybe.Value == 37).ToString(), "True");
