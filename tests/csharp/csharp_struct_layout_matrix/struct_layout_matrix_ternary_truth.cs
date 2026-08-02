// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
int seed = 113; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
