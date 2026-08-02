// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
var set = new System.Collections.Generic.HashSet<int>(); set.Add(27); set.Add(27); __Check((set.Count == 1).ToString(), "True");
