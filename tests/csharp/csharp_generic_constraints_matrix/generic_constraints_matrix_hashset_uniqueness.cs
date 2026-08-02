// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(80); set.Add(80); __Check((set.Count == 1).ToString(), "True");
