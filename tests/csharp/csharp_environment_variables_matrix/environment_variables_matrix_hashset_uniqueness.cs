// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(100); set.Add(100); __Check((set.Count == 1).ToString(), "True");
