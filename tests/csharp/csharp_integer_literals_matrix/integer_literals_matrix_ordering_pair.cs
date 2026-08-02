// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
int seed = 15; int right = seed + 1; __Check((seed < right).ToString(), "True");
