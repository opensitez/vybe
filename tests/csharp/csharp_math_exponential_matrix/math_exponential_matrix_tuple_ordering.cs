// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
var tuple = (left: 103, right: 104); __Check((tuple.left < tuple.right).ToString(), "True");
