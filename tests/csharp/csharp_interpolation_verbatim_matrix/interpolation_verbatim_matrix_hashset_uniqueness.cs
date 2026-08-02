// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(110); set.Add(110); __Check((set.Count == 1).ToString(), "True");
