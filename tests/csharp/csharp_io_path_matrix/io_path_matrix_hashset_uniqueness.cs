// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(122); set.Add(122); __Check((set.Count == 1).ToString(), "True");
