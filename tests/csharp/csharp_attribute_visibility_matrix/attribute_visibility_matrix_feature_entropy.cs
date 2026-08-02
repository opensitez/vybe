// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
string feature = "attribute_visibility_matrix:93"; __Check((feature.Length >= 1).ToString(), "True");
