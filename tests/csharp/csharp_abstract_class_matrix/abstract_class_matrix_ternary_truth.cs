// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
int seed = 72; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
