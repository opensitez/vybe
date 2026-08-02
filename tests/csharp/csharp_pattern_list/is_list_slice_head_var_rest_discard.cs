// vybe-test: csharp/csharp_pattern_list/is_list_slice_head_var_rest_discard
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{4,5,6}; if (data is [var first,..]) __Check((first).ToString(), "4");
