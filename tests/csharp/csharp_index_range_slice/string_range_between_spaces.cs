// vybe-test: csharp/csharp_index_range_slice/string_range_between_spaces
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="x y z"; __Check((text[2..4]).ToString(), "y ");
