// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
string feature = "delegate_invocation_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
