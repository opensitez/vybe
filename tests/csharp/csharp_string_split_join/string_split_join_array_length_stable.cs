// vybe-test: csharp/csharp_string_split_join/string_split_join_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
int seed = 21; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
