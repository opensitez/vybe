// vybe-test: csharp/csharp_expression_bodied_matrix/expression_bodied_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// expression_bodied_matrix
var tuple = (left: 106, right: 107); __Check((tuple.left < tuple.right).ToString(), "True");
