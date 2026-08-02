// vybe-test: csharp/csharp_null_conditional_matrix/null_conditional_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_null_conditional_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// null_conditional_matrix
string feature = "null_conditional_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
