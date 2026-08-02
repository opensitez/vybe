// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
int? maybe = 83; __Check((maybe.HasValue && maybe.Value == 83).ToString(), "True");
