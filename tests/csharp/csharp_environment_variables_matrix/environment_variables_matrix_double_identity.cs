// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
double seed = 100; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
