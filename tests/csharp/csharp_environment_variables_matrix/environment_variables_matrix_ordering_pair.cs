// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
int seed = 100; int right = seed + 1; __Check((seed < right).ToString(), "True");
