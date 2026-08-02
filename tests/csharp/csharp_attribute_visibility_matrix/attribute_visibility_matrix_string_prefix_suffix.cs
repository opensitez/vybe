// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
string feature = "attribute_visibility_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
