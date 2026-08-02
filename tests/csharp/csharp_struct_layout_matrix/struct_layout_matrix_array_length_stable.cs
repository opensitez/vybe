// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
int seed = 113; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
