// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(56); set.Add(56); __Check((set.Count == 1).ToString(), "True");
