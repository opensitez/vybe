// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
string feature = "environment_variables_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
