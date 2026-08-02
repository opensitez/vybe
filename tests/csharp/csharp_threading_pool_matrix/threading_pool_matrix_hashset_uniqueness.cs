// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(87); set.Add(87); __Check((set.Count == 1).ToString(), "True");
