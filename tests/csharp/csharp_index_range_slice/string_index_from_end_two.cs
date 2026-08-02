// vybe-test: csharp/csharp_index_range_slice/string_index_from_end_two
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="abcde"; __Check((text[^2]).ToString(), "100");
