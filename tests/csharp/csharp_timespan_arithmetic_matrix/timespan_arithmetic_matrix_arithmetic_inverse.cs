// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
int seed = 95; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
