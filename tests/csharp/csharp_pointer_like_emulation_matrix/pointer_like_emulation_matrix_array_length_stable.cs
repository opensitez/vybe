// vybe-test: csharp/csharp_pointer_like_emulation_matrix/pointer_like_emulation_matrix_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_pointer_like_emulation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pointer_like_emulation_matrix
int seed = 114; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
