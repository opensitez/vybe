// vybe-test: csharp/csharp_pattern_list/is_list_pair_discard_pattern_checks_length_only
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{9,1}; __Check((data is [_,_]).ToString(), "True");
