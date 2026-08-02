// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
var tuple = (left: 62, right: 63); __Check((tuple.left < tuple.right).ToString(), "True");
