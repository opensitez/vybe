// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
int seed = 105; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
