// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
int seed = 107; __Check((seed - seed == 0).ToString(), "True");
