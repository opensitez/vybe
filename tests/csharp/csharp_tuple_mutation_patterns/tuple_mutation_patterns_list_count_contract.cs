// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
var values = new System.Collections.Generic.List<int> { 37, 38, 37 }; __Check((values.Count == 3).ToString(), "True");
