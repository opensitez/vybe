// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(112); set.Add(112); __Check((set.Count == 1).ToString(), "True");
