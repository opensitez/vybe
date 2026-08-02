// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
int seed = 50; __Check((seed - seed == 0).ToString(), "True");
