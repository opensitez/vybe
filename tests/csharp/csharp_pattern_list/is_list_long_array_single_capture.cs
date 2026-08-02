// vybe-test: csharp/csharp_pattern_list/is_list_long_array_single_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long[] ids=new long[]{1000L}; if(ids is [var id]) __Check((id).ToString(), "1000");
