// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
int seed = 103; __Check((seed + 1 > seed).ToString(), "True");
