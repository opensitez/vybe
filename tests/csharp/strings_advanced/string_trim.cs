// vybe-test: csharp/strings_advanced/string_trim
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = "  hello  ";
__Check(("'" + s.Trim() + "'").ToString(), "'hello'");
__Check(("'" + s.TrimStart() + "'").ToString(), "'hello  '");
__Check(("'" + s.TrimEnd() + "'").ToString(), "'  hello'");
