// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
int seed = 103; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
