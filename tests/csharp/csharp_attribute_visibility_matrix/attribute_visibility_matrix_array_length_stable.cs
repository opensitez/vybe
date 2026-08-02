// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
int seed = 93; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
