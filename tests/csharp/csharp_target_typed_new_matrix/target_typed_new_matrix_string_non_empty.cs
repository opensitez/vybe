// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
string feature = "target_typed_new_matrix"; __Check((feature.Length > 0).ToString(), "True");
