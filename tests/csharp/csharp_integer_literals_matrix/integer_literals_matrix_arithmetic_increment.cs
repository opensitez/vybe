// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
int seed = 15; __Check((seed + 1 > seed).ToString(), "True");
