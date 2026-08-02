// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_mutation_patterns
var map = new System.Collections.Generic.Dictionary<int, int>(); map[37] = 38; __Check((map.ContainsKey(37) && map[37] == 38).ToString(), "True");
