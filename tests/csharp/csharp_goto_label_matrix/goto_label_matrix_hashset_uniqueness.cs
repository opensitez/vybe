// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(50); set.Add(50); __Check((set.Count == 1).ToString(), "True");
