// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
var tuple = (left: 110, right: 111); __Check((tuple.left < tuple.right).ToString(), "True");
