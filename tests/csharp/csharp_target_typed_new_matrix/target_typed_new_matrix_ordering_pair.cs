// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
int seed = 107; int right = seed + 1; __Check((seed < right).ToString(), "True");
