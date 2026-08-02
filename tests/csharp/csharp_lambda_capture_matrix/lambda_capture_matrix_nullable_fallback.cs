// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
int? maybe = null; int fallback = maybe ?? 75; __Check((fallback == 75).ToString(), "True");
