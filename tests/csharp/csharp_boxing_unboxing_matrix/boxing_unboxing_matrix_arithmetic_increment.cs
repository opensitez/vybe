// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
int seed = 62; __Check((seed + 1 > seed).ToString(), "True");
