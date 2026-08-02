// vybe-test: csharp/strings_advanced/string_toupper_tolower
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("Hello World".ToUpper()).ToString(), "HELLO WORLD");
__Check(("Hello World".ToLower()).ToString(), "hello world");
