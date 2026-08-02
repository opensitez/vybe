// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
string feature = "null_conditional_matrix:55"; __Check((feature.Length >= 1).ToString(), "True");
