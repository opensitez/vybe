// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
string feature = "constructor_null_guard_matrix"; __Check((feature.Length > 0).ToString(), "True");
