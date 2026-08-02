// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
string feature = "method_group_matrix:79"; __Check((feature.Length >= 1).ToString(), "True");
