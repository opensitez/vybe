// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
int seed = 50; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
