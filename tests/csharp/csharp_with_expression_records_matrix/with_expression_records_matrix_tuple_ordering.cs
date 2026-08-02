// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_records_matrix
var tuple = (left: 109, right: 110); __Check((tuple.left < tuple.right).ToString(), "True");
