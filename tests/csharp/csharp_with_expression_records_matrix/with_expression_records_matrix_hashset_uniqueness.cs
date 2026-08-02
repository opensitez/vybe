// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_records_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(109); set.Add(109); __Check((set.Count == 1).ToString(), "True");
