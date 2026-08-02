// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
int seed = 64; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
