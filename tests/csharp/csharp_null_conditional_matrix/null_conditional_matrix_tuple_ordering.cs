// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
var tuple = (left: 55, right: 56); __Check((tuple.left < tuple.right).ToString(), "True");
