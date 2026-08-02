// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
int seed = 45; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
