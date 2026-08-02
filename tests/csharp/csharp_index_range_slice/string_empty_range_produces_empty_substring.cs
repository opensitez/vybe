// vybe-test: csharp/csharp_index_range_slice/string_empty_range_produces_empty_substring
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="abc"; __Check((text[1..1].Length).ToString(), "0");
