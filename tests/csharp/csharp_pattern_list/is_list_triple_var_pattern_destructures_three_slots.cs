// vybe-test: csharp/csharp_pattern_list/is_list_triple_var_pattern_destructures_three_slots
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{1,2,3}; if (data is [var a,var b,var c]) __Check((a+b+c).ToString(), "6");
