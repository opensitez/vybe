// vybe-test: csharp/csharp_cast_runtime_matrix/cast_runtime_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_cast_runtime_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// cast_runtime_matrix
int seed = 61; __Check((seed + 1 > seed).ToString(), "True");
