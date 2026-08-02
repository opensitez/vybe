// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
var values = new System.Collections.Generic.List<int> { 75, 76, 75 }; __Check((values.Count == 3).ToString(), "True");
