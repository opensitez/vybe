// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(24); set.Add(24); __Check((set.Count == 1).ToString(), "True");
