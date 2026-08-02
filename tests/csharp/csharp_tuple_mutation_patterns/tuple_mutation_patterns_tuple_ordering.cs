// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
var tuple = (left: 37, right: 38); __Check((tuple.left < tuple.right).ToString(), "True");
