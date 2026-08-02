// vybe-test: csharp/csharp_pattern_list/is_list_slice_bookended_constants_match_middle_gap
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{1,2,3,4,5}; __Check((data is [1,..,5]).ToString(), "True");
