// vybe-test: csharp/csharp_cast_runtime_matrix/cast_runtime_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_cast_runtime_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// cast_runtime_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(61); set.Add(61); __Check((set.Count == 1).ToString(), "True");
