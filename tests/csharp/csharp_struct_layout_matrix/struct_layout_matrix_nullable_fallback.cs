// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
int? maybe = null; int fallback = maybe ?? 113; __Check((fallback == 113).ToString(), "True");
