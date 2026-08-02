// vybe-test: csharp/csharp_index_range_slice/string_single_char_range
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="dart"; __Check((text[1..2]).ToString(), "a");
