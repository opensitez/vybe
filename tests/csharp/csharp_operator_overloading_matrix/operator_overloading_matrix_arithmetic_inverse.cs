// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
int seed = 105; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
