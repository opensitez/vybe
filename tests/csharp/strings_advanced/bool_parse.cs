// vybe-test: csharp/strings_advanced/bool_parse
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool t = bool.Parse("True");
bool f = bool.Parse("False");
__Check((t).ToString(), "True");
__Check((f).ToString(), "False");
