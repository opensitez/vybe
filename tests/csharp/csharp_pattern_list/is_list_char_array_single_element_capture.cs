// vybe-test: csharp/csharp_pattern_list/is_list_char_array_single_element_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] chars=new[]{'x'}; if(chars is [var ch]) __Check((ch).ToString(), "x");
