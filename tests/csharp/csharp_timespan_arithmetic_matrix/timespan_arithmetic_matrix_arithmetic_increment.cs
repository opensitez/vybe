// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
int seed = 95; __Check((seed + 1 > seed).ToString(), "True");
