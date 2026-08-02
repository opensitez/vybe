// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
var set = new System.Collections.Generic.HashSet<int>(); set.Add(37); set.Add(37); __Check((set.Count == 1).ToString(), "True");
