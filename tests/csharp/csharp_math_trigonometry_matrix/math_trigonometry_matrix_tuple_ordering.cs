// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
var tuple = (left: 102, right: 103); __Check((tuple.left < tuple.right).ToString(), "True");
