// vybe-test: csharp/csharp_pattern_list/is_list_empty_pattern_rejects_non_empty_array
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{1}; __Check((data is []).ToString(), "False");
