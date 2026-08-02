// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(54); set.Add(54); __Check((set.Count == 1).ToString(), "True");
