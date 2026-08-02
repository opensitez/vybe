// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(79); set.Add(79); __Check((set.Count == 1).ToString(), "True");
