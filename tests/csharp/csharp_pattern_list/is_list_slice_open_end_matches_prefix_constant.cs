// vybe-test: csharp/csharp_pattern_list/is_list_slice_open_end_matches_prefix_constant
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{9,8,7}; __Check((data is [9,..]).ToString(), "True");
