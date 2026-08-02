// vybe-test: csharp/csharp_generic_variance_matrix/generic_variance_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_variance_matrix
var tuple = (left: 82, right: 83); __Check((tuple.left < tuple.right).ToString(), "True");
