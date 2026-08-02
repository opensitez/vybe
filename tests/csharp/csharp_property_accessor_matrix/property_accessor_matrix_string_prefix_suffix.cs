// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
string feature = "property_accessor_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
