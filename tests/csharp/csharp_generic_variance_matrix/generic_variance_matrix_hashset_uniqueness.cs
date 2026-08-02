// vybe-test: csharp/csharp_generic_variance_matrix/generic_variance_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_variance_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(82); set.Add(82); __Check((set.Count == 1).ToString(), "True");
