// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
var tuple = (left: 101, right: 102); __Check((tuple.left < tuple.right).ToString(), "True");
