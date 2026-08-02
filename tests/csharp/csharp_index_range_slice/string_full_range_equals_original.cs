// vybe-test: csharp/csharp_index_range_slice/string_full_range_equals_original
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="same"; __Check((text[..]==text).ToString(), "True");
