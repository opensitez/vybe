// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
int seed = 105; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
