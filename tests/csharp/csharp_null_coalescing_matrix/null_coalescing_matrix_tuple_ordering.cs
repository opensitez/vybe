// vybe-test: csharp/csharp_null_coalescing_matrix/null_coalescing_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_null_coalescing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_coalescing_matrix
var tuple = (left: 56, right: 57); __Check((tuple.left < tuple.right).ToString(), "True");
