// vybe-test: csharp/csharp_index_range_slice/string_range_unicode_characters
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="café"; __Check((text[1..3]).ToString(), "af");
