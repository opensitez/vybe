// vybe-test: csharp/csharp_pattern_list/is_list_first_constant_second_var_pattern
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{0,15}; if (data is [0,var tail]) __Check((tail).ToString(), "15");
