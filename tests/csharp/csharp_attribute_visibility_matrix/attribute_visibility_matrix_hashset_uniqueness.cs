// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(93); set.Add(93); __Check((set.Count == 1).ToString(), "True");
