// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
string feature = "attribute_visibility_matrix"; __Check((feature.Length > 0).ToString(), "True");
