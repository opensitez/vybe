// vybe-test: csharp/csharp_finally_cleanup_matrix/finally_cleanup_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_finally_cleanup_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// finally_cleanup_matrix
string feature = "finally_cleanup_matrix"; __Check((feature.Length > 0).ToString(), "True");
