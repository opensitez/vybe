// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
int seed = 121; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
