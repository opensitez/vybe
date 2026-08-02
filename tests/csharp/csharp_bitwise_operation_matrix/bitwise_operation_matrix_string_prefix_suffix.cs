// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
string feature = "bitwise_operation_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
