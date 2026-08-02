// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(94); set.Add(94); __Check((set.Count == 1).ToString(), "True");
