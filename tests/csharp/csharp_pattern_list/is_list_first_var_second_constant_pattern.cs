// vybe-test: csharp/csharp_pattern_list/is_list_first_var_second_constant_pattern
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{8,2}; if (data is [var head,2]) __Check((head).ToString(), "8");
