// vybe-test: csharp/csharp_index_range_slice/string_range_full_copy_via_open_bounds
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="abc"; __Check((text[..]).ToString(), "abc");
