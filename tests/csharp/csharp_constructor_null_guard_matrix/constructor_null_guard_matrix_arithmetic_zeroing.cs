// vybe-test: csharp/csharp_constructor_null_guard_matrix/constructor_null_guard_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_constructor_null_guard_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_null_guard_matrix
int seed = 126; __Check((seed - seed == 0).ToString(), "True");
