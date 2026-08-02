// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
int seed = 75; __Check((seed - seed == 0).ToString(), "True");
