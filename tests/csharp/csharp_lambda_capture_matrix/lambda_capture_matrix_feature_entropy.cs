// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
string feature = "lambda_capture_matrix:75"; __Check((feature.Length >= 1).ToString(), "True");
