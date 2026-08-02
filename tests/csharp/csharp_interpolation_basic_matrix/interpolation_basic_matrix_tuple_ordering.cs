// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
var tuple = (left: 112, right: 113); __Check((tuple.left < tuple.right).ToString(), "True");
