// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
var set = new System.Collections.Generic.HashSet<int>(); set.Add(28); set.Add(28); __Check((set.Count == 1).ToString(), "True");
