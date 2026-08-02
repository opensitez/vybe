// vybe-test: csharp/csharp_pattern_list/is_list_pair_var_pattern_binds_both_positions
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{3,4}; if (data is [var a,var b]) __Check((a+b).ToString(), "7");
