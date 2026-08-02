// vybe-test: csharp/csharp_index_range_slice/string_index_from_end_reads_last_char_code
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="xy"; __Check((text[^1]).ToString(), "121");
