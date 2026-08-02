// vybe-test: csharp/csharp_index_range_slice/string_range_from_end_indices
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="program"; __Check((text[^4..^1]).ToString(), "gra");
