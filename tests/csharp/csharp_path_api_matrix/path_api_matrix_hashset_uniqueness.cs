// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(123); set.Add(123); __Check((set.Count == 1).ToString(), "True");
