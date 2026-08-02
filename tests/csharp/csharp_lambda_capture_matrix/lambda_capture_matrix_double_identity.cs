// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
double seed = 75; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
