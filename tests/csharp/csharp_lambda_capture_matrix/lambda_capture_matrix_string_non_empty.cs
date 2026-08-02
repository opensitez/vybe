// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
string feature = "lambda_capture_matrix"; __Check((feature.Length > 0).ToString(), "True");
