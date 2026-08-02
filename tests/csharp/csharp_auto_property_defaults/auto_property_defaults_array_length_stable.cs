// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
int seed = 65; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
