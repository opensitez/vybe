// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(43); set.Add(43); __Check((set.Count == 1).ToString(), "True");
