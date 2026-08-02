// vybe-test: csharp/strings_advanced/char_isupper_islower
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((char.IsUpper('A')).ToString(), "True");
__Check((char.IsLower('a')).ToString(), "True");
__Check((char.IsDigit('5')).ToString(), "True");
__Check((char.IsLetter('x')).ToString(), "True");
__Check((char.IsWhiteSpace(' ')).ToString(), "True");
