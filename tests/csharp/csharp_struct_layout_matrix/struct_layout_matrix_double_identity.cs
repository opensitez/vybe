// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
double seed = 113; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
