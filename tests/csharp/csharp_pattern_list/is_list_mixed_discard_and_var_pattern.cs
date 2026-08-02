// vybe-test: csharp/csharp_pattern_list/is_list_mixed_discard_and_var_pattern
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{11,22,33}; if (data is [var a,_,var c]) __Check((a+c).ToString(), "44");
