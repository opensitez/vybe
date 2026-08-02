// vybe-test: csharp/csharp_constructor_chaining_matrix/constructor_chaining_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chaining_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_chaining_matrix
string feature = "constructor_chaining_matrix"; __Check((feature.Length > 0).ToString(), "True");
