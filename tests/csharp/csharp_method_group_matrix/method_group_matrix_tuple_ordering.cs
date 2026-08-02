// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
var tuple = (left: 79, right: 80); __Check((tuple.left < tuple.right).ToString(), "True");
