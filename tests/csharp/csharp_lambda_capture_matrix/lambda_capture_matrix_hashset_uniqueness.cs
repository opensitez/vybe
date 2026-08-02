// vybe-test: csharp/csharp_lambda_capture_matrix/lambda_capture_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_lambda_capture_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_capture_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(75); set.Add(75); __Check((set.Count == 1).ToString(), "True");
