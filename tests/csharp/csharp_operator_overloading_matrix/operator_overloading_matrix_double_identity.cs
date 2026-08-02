// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
double seed = 105; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
