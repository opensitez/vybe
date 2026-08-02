// vybe-test: csharp/csharp_raw_string_literals/raw_string_custom_delimiter_allows_unescaped_quotes
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text=""""""a "b" c""""""; __Check((text.Length>0).ToString(), "True");
