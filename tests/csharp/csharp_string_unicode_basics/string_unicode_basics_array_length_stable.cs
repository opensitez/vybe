// vybe-test: csharp/csharp_string_unicode_basics/string_unicode_basics_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_string_unicode_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_unicode_basics
int seed = 19; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
