// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
string feature = "integer_literals_matrix"; __Check((feature.Length > 0).ToString(), "True");
