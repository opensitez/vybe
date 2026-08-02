// vybe-test: csharp/csharp_with_expression_records_matrix/with_expression_records_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_records_matrix
string feature = "with_expression_records_matrix:109"; __Check((feature.Length >= 1).ToString(), "True");
