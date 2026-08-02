// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(119); set.Add(119); __Check((set.Count == 1).ToString(), "True");
