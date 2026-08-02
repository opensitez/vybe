// vybe-test: csharp/csharp_index_range_slice/string_slice_does_not_mutate_original
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="keep"; var part=text[1..3]; __Check((text).ToString(), "keep"); __Check((part).ToString(), "ee");
