// vybe-test: csharp/csharp_pattern_list/is_list_pair_constant_pattern_matches_exact_sequence
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{1,2}; __Check((data is [1,2]).ToString(), "True");
