// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
int seed = 80; __Check((seed - seed == 0).ToString(), "True");
