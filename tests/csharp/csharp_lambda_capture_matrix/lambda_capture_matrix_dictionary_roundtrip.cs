// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[75] = 76; __Check((map.ContainsKey(75) && map[75] == 76).ToString(), "True");
