// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
int seed = 119; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
