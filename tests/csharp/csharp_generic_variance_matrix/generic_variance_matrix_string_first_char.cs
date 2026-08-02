// vybe-test: csharp/csharp_generic_variance_matrix/generic_variance_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_variance_matrix
string feature = "generic_variance_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
