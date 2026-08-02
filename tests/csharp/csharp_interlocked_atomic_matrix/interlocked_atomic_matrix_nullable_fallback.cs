// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
int? maybe = null; int fallback = maybe ?? 83; __Check((fallback == 83).ToString(), "True");
