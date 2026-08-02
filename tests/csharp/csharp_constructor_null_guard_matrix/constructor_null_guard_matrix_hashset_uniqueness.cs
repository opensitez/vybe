// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(126); set.Add(126); __Check((set.Count == 1).ToString(), "True");
