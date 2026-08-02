// vybe-test: csharp/csharp_pattern_list/switch_statement_list_pattern_case_with_var_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data=new[]{3,9}; string tag=""; switch(data){case[var a,var b]:tag=(a+b).ToString();break;default:tag="0";break;} __Check((tag).ToString(), "12");
