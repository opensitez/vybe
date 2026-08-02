// vybe-test: csharp/csharp_cast_runtime_matrix/cast_runtime_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_cast_runtime_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// cast_runtime_matrix
string feature = "cast_runtime_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
