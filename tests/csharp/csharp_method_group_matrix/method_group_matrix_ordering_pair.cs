// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
int seed = 79; int right = seed + 1; __Check((seed < right).ToString(), "True");
