// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
int seed = 26; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
