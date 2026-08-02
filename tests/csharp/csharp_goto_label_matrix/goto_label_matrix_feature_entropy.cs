// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
string feature = "goto_label_matrix:50"; __Check((feature.Length >= 1).ToString(), "True");
