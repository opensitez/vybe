// vybe-test: csharp/csharp_pattern_list/is_list_empty_pattern_matches_zero_length_array
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new int[]{}; __Check((data is []).ToString(), "True");
