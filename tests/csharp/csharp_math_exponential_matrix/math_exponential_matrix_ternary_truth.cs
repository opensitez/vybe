// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
int seed = 103; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
