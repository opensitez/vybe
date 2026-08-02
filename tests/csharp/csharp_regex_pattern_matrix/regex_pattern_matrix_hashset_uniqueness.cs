// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(99); set.Add(99); __Check((set.Count == 1).ToString(), "True");
