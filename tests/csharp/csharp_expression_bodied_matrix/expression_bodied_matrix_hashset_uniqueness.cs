// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(106); set.Add(106); __Check((set.Count == 1).ToString(), "True");
