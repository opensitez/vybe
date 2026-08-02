// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(98); set.Add(98); __Check((set.Count == 1).ToString(), "True");
