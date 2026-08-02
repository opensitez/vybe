// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
var tuple = (left: 75, right: 76); __Check((tuple.left < tuple.right).ToString(), "True");
