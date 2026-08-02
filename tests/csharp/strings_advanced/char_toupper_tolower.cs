// vybe-test: csharp/strings_advanced/char_toupper_tolower
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((char.ToUpper('a')).ToString(), "A");
__Check((char.ToLower('Z')).ToString(), "z");
