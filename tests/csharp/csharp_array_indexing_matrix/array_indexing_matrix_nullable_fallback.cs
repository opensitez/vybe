// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
int? maybe = null; int fallback = maybe ?? 24; __Check((fallback == 24).ToString(), "True");
