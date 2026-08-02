// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// do_while_matrix
string feature = "do_while_matrix"; __Check((feature.Length > 0).ToString(), "True");
