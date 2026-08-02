// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
string feature = "math_exponential_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
