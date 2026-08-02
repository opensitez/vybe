// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
string feature = "abstract_class_matrix:72"; __Check((feature.Length >= 1).ToString(), "True");
