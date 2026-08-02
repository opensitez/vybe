// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
int seed = 75; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
