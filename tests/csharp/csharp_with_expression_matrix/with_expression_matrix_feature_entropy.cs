// vybe-test: csharp/csharp_with_expression_matrix/with_expression_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// with_expression_matrix
string feature = "with_expression_matrix:108"; __Check((feature.Length >= 1).ToString(), "True");
