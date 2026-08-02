// vybe-test: csharp/csharp_index_range_slice/string_range_single_char_at_start
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="open"; __Check((text[0..1]).ToString(), "o");
