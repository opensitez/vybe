// vybe-test: csharp/strings_advanced/verbatim_string_with_quotes
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = @"He said ""hello""";
__Check((s).ToString(), "He said \"hello\"");
