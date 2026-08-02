// vybe-test: csharp/csharp_index_range_slice/index_from_end_on_string_with_spaces
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="a b c"; __Check((text[^2]).ToString(), "32");
