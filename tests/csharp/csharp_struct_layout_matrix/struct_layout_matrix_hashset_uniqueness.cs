// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(113); set.Add(113); __Check((set.Count == 1).ToString(), "True");
