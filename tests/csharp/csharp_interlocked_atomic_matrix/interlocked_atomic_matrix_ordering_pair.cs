// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
int seed = 83; int right = seed + 1; __Check((seed < right).ToString(), "True");
