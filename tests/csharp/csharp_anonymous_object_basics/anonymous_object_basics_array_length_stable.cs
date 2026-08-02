// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
int seed = 38; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
