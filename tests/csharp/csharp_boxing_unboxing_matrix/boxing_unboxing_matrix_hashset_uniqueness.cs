// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(62); set.Add(62); __Check((set.Count == 1).ToString(), "True");
