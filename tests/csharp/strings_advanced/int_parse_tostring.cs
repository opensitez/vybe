// vybe-test: csharp/strings_advanced/int_parse_tostring
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = int.Parse("42");
__Check((x + 8).ToString(), "50");
__Check((x.ToString()).ToString(), "42");
