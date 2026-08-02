// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
string feature = "null_conditional_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
