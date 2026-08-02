// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(64); set.Add(64); __Check((set.Count == 1).ToString(), "True");
