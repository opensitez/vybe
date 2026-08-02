// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(72); set.Add(72); __Check((set.Count == 1).ToString(), "True");
