// vybe-test: csharp/csharp_serialization_json_surface/serialization_json_surface_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_serialization_json_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// serialization_json_surface
int seed = 91; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
