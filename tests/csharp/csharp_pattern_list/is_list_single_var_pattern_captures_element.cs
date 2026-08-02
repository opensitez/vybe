// vybe-test: csharp/csharp_pattern_list/is_list_single_var_pattern_captures_element
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{42}; if (data is [var n]) __Check((n).ToString(), "42");
