// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
int seed = 95; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
