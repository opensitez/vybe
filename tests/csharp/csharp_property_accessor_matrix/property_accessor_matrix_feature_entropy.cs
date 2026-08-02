// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
string feature = "property_accessor_matrix:64"; __Check((feature.Length >= 1).ToString(), "True");
