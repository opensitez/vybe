// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
int seed = 100; __Check((seed + 1 > seed).ToString(), "True");
