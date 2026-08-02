// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
string feature = "lambda_capture_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
