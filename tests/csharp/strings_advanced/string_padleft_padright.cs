// vybe-test: csharp/strings_advanced/string_padleft_padright
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "hi";
__Check(("'" + s.PadLeft(6) + "'").ToString(), "'    hi'");
__Check(("'" + s.PadRight(6) + "'").ToString(), "'hi    '");
__Check(("'" + s.PadLeft(6, '*') + "'").ToString(), "'****hi'");
