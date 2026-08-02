// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
int seed = 107; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
