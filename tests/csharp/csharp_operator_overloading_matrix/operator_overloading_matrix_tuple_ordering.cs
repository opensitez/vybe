// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
var tuple = (left: 105, right: 106); __Check((tuple.left < tuple.right).ToString(), "True");
