// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
int seed = 126; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
