// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
int seed = 29; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
