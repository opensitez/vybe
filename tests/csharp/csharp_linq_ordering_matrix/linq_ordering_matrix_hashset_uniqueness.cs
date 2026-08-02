// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(121); set.Add(121); __Check((set.Count == 1).ToString(), "True");
