// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
int seed = 15; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
