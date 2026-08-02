// vybe-test: csharp/csharp_extension_method_matrix/extension_method_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_extension_method_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// extension_method_matrix
var tuple = (left: 78, right: 79); __Check((tuple.left < tuple.right).ToString(), "True");
