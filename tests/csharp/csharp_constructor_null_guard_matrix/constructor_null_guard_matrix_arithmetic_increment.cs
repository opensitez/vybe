// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
int seed = 126; __Check((seed + 1 > seed).ToString(), "True");
