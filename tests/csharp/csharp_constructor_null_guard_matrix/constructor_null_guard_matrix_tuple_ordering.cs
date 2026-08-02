// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
var tuple = (left: 126, right: 127); __Check((tuple.left < tuple.right).ToString(), "True");
