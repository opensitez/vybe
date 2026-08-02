// vybe-test: csharp/csharp_conversion_builtins_matrix/conversion_builtins_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_conversion_builtins_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// conversion_builtins_matrix
var tuple = (left: 124, right: 125); __Check((tuple.left < tuple.right).ToString(), "True");
