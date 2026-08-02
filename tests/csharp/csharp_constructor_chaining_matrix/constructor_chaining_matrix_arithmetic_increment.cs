// vybe-test: csharp/csharp_constructor_chaining_matrix/constructor_chaining_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chaining_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_chaining_matrix
int seed = 68; __Check((seed + 1 > seed).ToString(), "True");
