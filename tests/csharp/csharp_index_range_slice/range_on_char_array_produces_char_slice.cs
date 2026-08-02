// vybe-test: csharp/csharp_index_range_slice/range_on_char_array_produces_char_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] letters={'a','b','c','d'}; var slice=letters[1..3]; __Check((slice.Length).ToString(), "2"); __Check((slice[0]).ToString(), "98");
