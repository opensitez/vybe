// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(83); set.Add(83); __Check((set.Count == 1).ToString(), "True");
