// vybe-test: csharp/csharp_constructor_overload_surface/constructor_overload_surface_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_constructor_overload_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// constructor_overload_surface
int seed = 67; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
