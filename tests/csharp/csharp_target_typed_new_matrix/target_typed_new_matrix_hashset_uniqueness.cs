// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(107); set.Add(107); __Check((set.Count == 1).ToString(), "True");
