// vybe-test: csharp/csharp_switch_expression_matrix/switch_expression_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// switch_expression_matrix
var tuple = (left: 43, right: 44); __Check((tuple.left < tuple.right).ToString(), "True");
