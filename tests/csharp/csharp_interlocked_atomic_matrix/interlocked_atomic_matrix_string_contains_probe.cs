// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
string feature = "interlocked_atomic_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
