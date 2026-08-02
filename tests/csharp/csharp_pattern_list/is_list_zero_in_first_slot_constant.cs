// vybe-test: csharp/csharp_pattern_list/is_list_zero_in_first_slot_constant
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data=new[]{0,42}; if(data is [0,var v]) __Check((v).ToString(), "42");
