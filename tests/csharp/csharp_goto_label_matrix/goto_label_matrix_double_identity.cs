// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
double seed = 50; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
