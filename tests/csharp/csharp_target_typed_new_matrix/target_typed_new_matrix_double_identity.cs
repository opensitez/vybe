// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
double seed = 107; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
