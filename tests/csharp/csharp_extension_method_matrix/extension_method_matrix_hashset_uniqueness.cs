// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(78); set.Add(78); __Check((set.Count == 1).ToString(), "True");
