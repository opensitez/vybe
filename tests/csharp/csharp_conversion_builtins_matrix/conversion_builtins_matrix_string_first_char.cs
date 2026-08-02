// vybe-test: csharp/csharp_conversion_builtins_matrix/conversion_builtins_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_conversion_builtins_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// conversion_builtins_matrix
string feature = "conversion_builtins_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
