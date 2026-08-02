// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
double seed = 95; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
