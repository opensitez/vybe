// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
string feature = "method_group_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
