// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
var tuple = (left: 81, right: 82); __Check((tuple.left < tuple.right).ToString(), "True");
