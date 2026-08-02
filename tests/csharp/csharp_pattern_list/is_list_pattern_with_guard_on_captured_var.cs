// vybe-test: csharp/csharp_pattern_list/is_list_pattern_with_guard_on_captured_var
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data=new[]{4,8}; if(data is [var a,var b] && a<b) __Check(("ordered").ToString(), "ordered");
