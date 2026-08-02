// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
string feature = "operator_overloading_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
