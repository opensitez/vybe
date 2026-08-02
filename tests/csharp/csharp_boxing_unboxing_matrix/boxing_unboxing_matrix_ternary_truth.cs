// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
int seed = 62; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
