// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
var tuple = (left: 54, right: 55); __Check((tuple.left < tuple.right).ToString(), "True");
