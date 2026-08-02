// vybe-test: csharp/csharp_comparison_operators_surface/comparison_operators_surface_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_comparison_operators_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// comparison_operators_surface
int seed = 13; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
