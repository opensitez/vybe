// vybe-test: csharp/csharp_pattern_list/is_list_single_discard_pattern_accepts_one_element
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{7}; __Check((data is [_]).ToString(), "True");
